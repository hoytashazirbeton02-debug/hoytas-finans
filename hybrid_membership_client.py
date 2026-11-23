"""Use the central membership endpoints without exposing relay credentials."""
import json
import re
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import Request, build_opener

from hybrid_http_relay import NoRedirect
from hybrid_enrollment import MAX_BASELINE_BYTES


class MembershipClient:
    def __init__(self, url, *, allow_loopback_http=False):
        endpoint = urlsplit(url)
        local = allow_loopback_http and endpoint.scheme == "http" and endpoint.hostname == "127.0.0.1"
        if (endpoint.scheme != "https" and not local) or not endpoint.hostname or endpoint.username or endpoint.password or endpoint.query or endpoint.fragment:
            raise ValueError("Geçerli bir HTTPS üyelik adresi gerekli.")
        self.url = url.rstrip("/")
        self.opener = build_opener(NoRedirect())

    def warm_up(self):
        """Wake the service with an anonymous request while the form is open."""
        try:
            with self.opener.open(Request(self.url + "/"), timeout=75):
                pass
        except HTTPError as error:
            error.close()
        except OSError:
            pass

    def request(self, username, password, invitation):
        return self._call("request", {"username": username, "password": password}, invitation)

    def setup_admin(self, username, password, code):
        return self._call("setup", {"username": username, "password": password}, code)

    def login(self, username, password):
        account = self._call("login", {"username": username, "password": password})
        if (not isinstance(account.get("session"), str)
                or not re.fullmatch(r"[A-Za-z0-9_-]{43}", account["session"])
                or account.get("role") not in ("admin", "member")
                or not isinstance(account.get("username"), str)
                or not account["username"]):
            raise RuntimeError("Sunucu geçerli bir giriş yanıtı vermedi.")
        return account

    def pending(self, session):
        return self._call("pending", token=session)["pending"]

    def approve(self, session, username):
        return self._call("approve", {"username": username}, session)

    def publish_baseline(self, session, baseline):
        return self._call("baseline", baseline, session)

    def enroll(self, session, device):
        return self._call("enroll", {"device": device}, session)

    def _call(self, path, body=None, token=None):
        headers = {"Content-Type": "application/json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        payload = json.dumps(body, ensure_ascii=False, allow_nan=False, separators=(",", ":")).encode() if body is not None else None
        if payload is not None and len(payload) > MAX_BASELINE_BYTES:
            raise ValueError("Başlangıç kayıtları tek aktarım için fazla büyük.")
        command = Request(self.url + "/membership/" + path, data=payload, headers=headers)
        try:
            limit = MAX_BASELINE_BYTES + 1024 if path == "enroll" else 1024 * 1024
            # Free hosting may need roughly a minute to start after being idle.
            with self.opener.open(command, timeout=75 if path == "login" else 20) as response:
                raw = response.read(limit + 1)
            if len(raw) > limit:
                raise RuntimeError("Sunucu yanıtı izin verilen boyutu aştı.")
            try:
                result = json.loads(raw)
            except ValueError:
                raise RuntimeError("Sunucu hazır olmayabilir. Biraz bekleyip tekrar deneyin.") from None
            if not isinstance(result, dict):
                raise RuntimeError("Sunucu geçerli bir üyelik yanıtı vermedi.")
            return result
        except HTTPError as error:
            with error:
                code = error.code
            messages = {400: "Bilgileri kontrol edin. Bu kullanıcı adı alınmış olabilir.",
                        403: "Erişim reddedildi. Giriş bilgilerini, daveti veya onay durumunu kontrol edin.",
                        404: "Bu sunucuda merkezi üyelik açılmamış.",
                        409: "Bu başvuru artık onay beklemiyor.",
                        429: "Çok fazla deneme yapıldı. Birkaç dakika sonra tekrar deneyin."}
            if code == 403 and path == "setup":
                messages[403] = "Kurulum kodu geçersiz veya yönetici hesabı zaten oluşturulmuş."
            if code == 409 and path == "baseline":
                messages[409] = "Sunucuyla senkronizasyonu bekleyip tekrar paylaşın."
            if code == 409 and path == "enroll":
                messages[409] = "Yönetici henüz başlangıç kayıtlarını paylaşmadı."
            raise RuntimeError(messages.get(code, "Üyelik servisine şu anda ulaşılamıyor.")) from None
        except (URLError, TimeoutError, ConnectionError):
            raise RuntimeError("Merkezi sunucuya bağlanılamadı. İnternet bağlantınızı kontrol edip tekrar deneyin.") from None
