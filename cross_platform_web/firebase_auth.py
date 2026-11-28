"""Firebase Authentication helper (optional, config-driven).

Adds Firebase email sign-up/sign-in (with
email verification) to the app's login and signup pages, and verifies
Firebase ID tokens on the server with firebase-admin. The app works exactly as before when Firebase is NOT
configured -- the Firebase buttons simply do not appear.

Configuration (no secrets are ever hardcoded here):
  1. FIREBASE_WEB_CONFIG environment variable -- the JSON object shown in
     Firebase console > Project settings > Your apps > SDK setup and
     configuration (apiKey, authDomain, projectId, storageBucket,
     messagingSenderId, appId). Example:
         export FIREBASE_WEB_CONFIG='{"apiKey":"...","authDomain":"...","projectId":"...","storageBucket":"...","messagingSenderId":"...","appId":"..."}'
  2. OR a firebase-config.json file next to this module, with the same
     JSON object. (Do NOT commit it to a public repo if it ever holds
     anything sensitive; the web config itself is a public client
     identifier, but keep the file out of git anyway.)

Server-side token verification uses PyJWT against Google's published
public keys: verifying an ID token only needs the RS256 signature check
plus audience / issuer / expiry, so no service account, no private key
and no Google credentials are created, stored or needed here.
(firebase-admin cannot do this on Render: its Auth client resolves
Application Default Credentials when it is constructed, so
verify_id_token fails with DefaultCredentialsError wherever no service
account exists -- verified 29 Nov 2026 on all four apps. A service
account is only required for Admin SDK calls that write to Firebase,
which this app does not use.)
"""

import json
import os
import time
import urllib.request

_CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "firebase-config.json"
)
# Pinned Firebase JS SDK (compat build) used by the login page.
_FIREBASE_JS_VERSION = "10.12.2"

_cache = {"loaded": False, "config": None}


class FirebaseNotConfigured(Exception):
    """Firebase has no web config, or PyJWT is not installed."""


_GOOGLE_JWKS_URL = (
    "https://www.googleapis.com/service_accounts/v1/jwk/"
    "securetoken@system.gserviceaccount.com"
)
_JWKS_CACHE_SECONDS = 3600
_jwks = {"fetched_at": 0.0, "keys": {}}


def _fetch_jwks():
    """Fetch Google's public keys for Firebase ID tokens (cached 1h).

    Raises FirebaseNotConfigured when the keys cannot be fetched
    (network problem on the server); callers map that to HTTP 503.
    """
    now = time.time()
    if _jwks["keys"] and now - _jwks["fetched_at"] < _JWKS_CACHE_SECONDS:
        return _jwks["keys"]
    req = urllib.request.Request(
        _GOOGLE_JWKS_URL, headers={"User-Agent": "python-urllib"}
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.load(resp)
    except Exception as exc:
        raise FirebaseNotConfigured(
            "Firebase sign-in cannot be verified right now: Google's "
            "public keys could not be fetched (%s). Please try again." % (exc,)
        ) from exc
    keys = {}
    for entry in data.get("keys", []):
        if entry.get("kid"):
            keys[entry["kid"]] = entry
    if not keys:
        raise FirebaseNotConfigured(
            "Firebase sign-in cannot be verified right now: Google "
            "returned no public keys. Please try again."
        )
    _jwks["fetched_at"] = now
    _jwks["keys"] = keys
    return keys


def _signing_key(kid):
    """RS256 public key for the token's kid; refetches once on a miss."""
    import jwt

    keys = _fetch_jwks()
    entry = keys.get(kid) if kid else None
    if entry is None:
        _jwks["fetched_at"] = 0.0
        keys = _fetch_jwks()
        entry = keys.get(kid) if kid else None
    if entry is None:
        raise InvalidFirebaseToken(
            "Invalid or expired Firebase sign-in. Please sign in again."
        )
    return jwt.algorithms.RSAAlgorithm.from_jwk(json.dumps(entry))


class InvalidFirebaseToken(Exception):
    """A Firebase ID token is missing, bad, expired or revoked."""


def _load_config():
    """Read the Firebase web config (env var first, then local file).

    Returns the config dict, or None when nothing usable is configured.
    A config counts only when it has at least apiKey and projectId.
    """
    if _cache["loaded"]:
        return _cache["config"]
    cfg = None
    raw = os.environ.get("FIREBASE_WEB_CONFIG", "").strip()
    if raw:
        try:
            cfg = json.loads(raw)
        except ValueError:
            cfg = None
    if cfg is None and os.path.exists(_CONFIG_FILE):
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as fh:
                cfg = json.load(fh)
        except (ValueError, OSError):
            cfg = None
    if not isinstance(cfg, dict) or not cfg.get("apiKey") or not cfg.get("projectId"):
        cfg = None
    _cache["loaded"] = True
    _cache["config"] = cfg
    return cfg


def is_configured():
    """True when a usable Firebase web config is present."""
    return _load_config() is not None


def web_config():
    """The Firebase web config dict, or None when not configured."""
    return _load_config()


def web_config_json():
    """The Firebase web config as compact JSON for the login page.

    Returns None when Firebase is not configured. The JSON is escaped
    so it can be embedded inside a <script> block safely.
    """
    cfg = _load_config()
    if cfg is None:
        return None
    return json.dumps(cfg, separators=(",", ":")).replace("<", "\\u003c")


def verify_id_token(id_token):
    """Verify a Firebase ID token; return the decoded claims dict.

    Verifies the RS256 signature against Google's published public keys
    and checks audience (this project's id), issuer, expiry and subject.
    No Google credentials are needed.

    Raises FirebaseNotConfigured (caller maps to HTTP 503) or
    InvalidFirebaseToken (caller maps to HTTP 401).
    """
    if not id_token or not isinstance(id_token, str):
        raise InvalidFirebaseToken("Missing ID token.")
    cfg = _load_config()
    if cfg is None:
        raise FirebaseNotConfigured(
            "Firebase is not configured on this server yet. Set the "
            "FIREBASE_WEB_CONFIG environment variable (or add "
            "firebase-config.json next to firebase_auth.py) with the "
            "Firebase web config from Project settings > Your apps, "
            "then restart the app."
        )
    try:
        import jwt
    except ImportError as exc:
        raise FirebaseNotConfigured(
            "Firebase is not set up on this server: the PyJWT "
            "package is not installed. Run: pip install -r requirements.txt"
        ) from exc
    project_id = cfg["projectId"]
    try:
        kid = jwt.get_unverified_header(id_token).get("kid")
    except Exception as exc:
        raise InvalidFirebaseToken(
            "Invalid or expired Firebase sign-in. Please sign in again."
        ) from exc
    key = _signing_key(kid)
    try:
        claims = jwt.decode(
            id_token,
            key,
            algorithms=["RS256"],
            audience=project_id,
            issuer="https://securetoken.google.com/" + project_id,
            options={"require": ["exp", "iat", "aud", "iss", "sub"]},
        )
    except Exception as exc:
        raise InvalidFirebaseToken(
            "Invalid or expired Firebase sign-in. Please sign in again."
        ) from exc
    if not isinstance(claims, dict) or not claims.get("sub"):
        raise InvalidFirebaseToken(
            "Invalid or expired Firebase sign-in. Please sign in again."
        )
    # firebase-admin returns the subject as "uid"; callers use that name.
    claims["uid"] = claims["sub"]
    return claims


def login_html():
    """HTML block for the login/signup pages: Firebase email auth.

    Returns "" when Firebase is not configured, so the pages render
    exactly as before (no scripts loaded, no JS errors). When configured,
    returns a divider, a Firebase
    sign-up form and a Firebase sign-in form, plus the Firebase JS SDK
    (compat build).

    Sign-up flow (client): createUserWithEmailAndPassword, then
    sendEmailVerification. The user is NOT logged in to the Flask app
    at this point -- they must click the confirmation link in the email
    and then sign in; the server also rejects ID tokens whose
    email_verified claim is false. Sign-in flow:
    signInWithEmailAndPassword; when the account's email is not
    verified, a message and a resend-confirmation button are shown and
    no token is posted to /auth/firebase.
    """
    cfg_json = web_config_json()
    if cfg_json is None:
        return ""
    v = _FIREBASE_JS_VERSION
    return """
<div class="firebase-auth" style="margin-top:18px;border-top:1px solid #dbe2ea;padding-top:16px">
  <p class="muted" style="text-align:center;margin:0 0 10px">or continue with Firebase</p>
  <form id="firebase-signup-form">
    <p style="margin:12px 0 4px;font-weight:600">Create a Firebase account</p>
    <label>Email<input name="fb-signup-email" type="email" required placeholder="you@example.com"></label>
    <label>Password<input name="fb-signup-password" type="password" required minlength="6" placeholder="At least 6 characters"></label>
    <p><button type="submit">Create account with Firebase</button></p>
  </form>
  <form id="firebase-login-form">
    <p style="margin:12px 0 4px;font-weight:600">Sign in with Firebase</p>
    <label>Email<input name="fb-login-email" type="email" required placeholder="you@example.com"></label>
    <label>Password<input name="fb-login-password" type="password" required placeholder="Your Firebase password"></label>
    <p><button type="submit">Sign in with Firebase</button></p>
  </form>
  <p><button type="button" id="firebase-resend-btn" style="display:none">Resend confirmation email</button></p>
  <p class="muted" id="firebase-info" style="color:#067647;display:none"></p>
  <p class="muted" id="firebase-error" style="color:#b42318;display:none"></p>
  <script src="https://www.gstatic.com/firebasejs/""" + v + """/firebase-app-compat.js"></script>
  <script src="https://www.gstatic.com/firebasejs/""" + v + """/firebase-auth-compat.js"></script>
  <script>
  (function(){
    var cfg = """ + cfg_json + """;
    if (!window.firebase || !cfg || !cfg.apiKey) return;
    try { firebase.initializeApp(cfg); } catch (e) { if (String(e && e.message).indexOf("already exists") === -1) return; }
    var auth = firebase.auth();
    var currentUser = null;
    function showError(msg){ var el = document.getElementById("firebase-error"); el.style.display = "block"; el.textContent = msg; }
    function showInfo(msg){ var el = document.getElementById("firebase-info"); el.style.display = "block"; el.textContent = msg; }
    function hideMsgs(){ document.getElementById("firebase-error").style.display = "none"; document.getElementById("firebase-info").style.display = "none"; }
    function sendToken(user){
      return user.getIdToken().then(function(token){
        return fetch("/auth/firebase", {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify({idToken: token})});
      }).then(function(r){ return r.json().then(function(j){ return {status: r.status, body: j}; }); })
      .then(function(res){
        if (res.status === 200 && res.body && res.body.ok) { window.location.href = "/"; }
        else { showError((res.body && res.body.error) || "Firebase sign-in failed."); }
      }).catch(function(err){ showError("Firebase sign-in failed: " + err.message); });
    }
    document.getElementById("firebase-signup-form").addEventListener("submit", function(ev){
      ev.preventDefault();
      hideMsgs();
      var f = ev.target;
      auth.createUserWithEmailAndPassword(f["fb-signup-email"].value, f["fb-signup-password"].value)
        .then(function(cred){
          return cred.user.sendEmailVerification().then(function(){
            currentUser = null;
            return auth.signOut();
          }).then(function(){
            showInfo("Confirmation link sent to " + f["fb-signup-email"].value + ". Click the link in your email, then sign in below.");
          });
        })
        .catch(function(err){ showError(err.message); });
    });
    document.getElementById("firebase-login-form").addEventListener("submit", function(ev){
      ev.preventDefault();
      hideMsgs();
      var f = ev.target;
      auth.signInWithEmailAndPassword(f["fb-login-email"].value, f["fb-login-password"].value)
        .then(function(cred){
          var user = cred.user;
          if (!user.emailVerified) {
            currentUser = user;
            document.getElementById("firebase-resend-btn").style.display = "";
            showError("Please confirm your email first. We sent a confirmation link to " + user.email + ".");
            return null;
          }
          return sendToken(user);
        })
        .catch(function(err){ showError(err.message); });
    });
    document.getElementById("firebase-resend-btn").addEventListener("click", function(){
      hideMsgs();
      if (!currentUser) { showError("Please sign in with your Firebase email and password first."); return; }
      currentUser.sendEmailVerification().then(function(){
        showInfo("Confirmation link sent again to " + currentUser.email + ".");
      }).catch(function(err){ showError(err.message); });
    });
  })();
  </script>
</div>
"""
