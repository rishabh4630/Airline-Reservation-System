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

Server-side token verification uses firebase-admin with NO service
account: verifying an ID token only needs Google's public keys and the
project id, so no private key is created, stored or needed here. A
service account is only required for Admin SDK calls that write to
Firebase (not used by this app).
"""

import json
import os

_CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "firebase-config.json"
)
# Pinned Firebase JS SDK (compat build) used by the login page.
_FIREBASE_JS_VERSION = "10.12.2"

_cache = {"loaded": False, "config": None}


class FirebaseNotConfigured(Exception):
    """Firebase has no web config, or firebase-admin is not installed."""


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


def _admin_auth():
    """Return the firebase_admin.auth module, ready to verify ID tokens.

    Initialises the default Firebase app with only the project id from
    the web config. No credential / service account is used: ID-token
    verification needs only Google's public keys, which firebase-admin
    fetches itself.
    """
    try:
        import firebase_admin
        from firebase_admin import auth as _auth
    except ImportError as exc:
        raise FirebaseNotConfigured(
            "Firebase is not set up on this server: the firebase-admin "
            "package is not installed. Run: pip install -r requirements.txt"
        ) from exc
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
        firebase_admin.get_app()
    except ValueError:
        try:
            firebase_admin.initialize_app(options={"projectId": cfg["projectId"]})
        except Exception as exc:  # e.g. no credentials found
            raise FirebaseNotConfigured(
                "Firebase could not be initialised on this server: %s" % (exc,)
            ) from exc
    return _auth


def verify_id_token(id_token):
    """Verify a Firebase ID token; return the decoded claims dict.

    Raises FirebaseNotConfigured (caller maps to HTTP 503) or
    InvalidFirebaseToken (caller maps to HTTP 401).
    """
    if not id_token or not isinstance(id_token, str):
        raise InvalidFirebaseToken("Missing ID token.")
    auth = _admin_auth()
    try:
        claims = auth.verify_id_token(id_token, check_revoked=False)
    except FirebaseNotConfigured:
        raise
    except Exception as exc:
        # firebase-admin raises InvalidIdTokenError / ExpiredIdTokenError /
        # RevokedIdTokenError (and CertificateFetchError) for bad tokens.
        raise InvalidFirebaseToken(
            "Invalid or expired Firebase sign-in. Please sign in again."
        ) from exc
    if not isinstance(claims, dict) or not claims.get("uid"):
        raise InvalidFirebaseToken(
            "Invalid or expired Firebase sign-in. Please sign in again."
        )
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
