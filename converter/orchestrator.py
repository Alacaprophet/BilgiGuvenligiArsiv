"""Yuksek seviyeli donusum yonlendiricisi.

Arayuz (GUI) yalnizca bu modulu cagirir; hangi motorun kullanilacagina burasi
karar verir. Tum fonksiyonlar ``progress(mesaj, oran)`` geri cagrimi alir
(oran 0..1, bilinmiyorsa -1).
"""

from __future__ import annotations

import os
import shutil
import tempfile
from typing import Callable, Dict, List, Optional

from . import engine_libpff, engine_outlook

Progress = Optional[Callable[[str, float], None]]


class ConversionError(Exception):
    """Donusum sirasinda olusan, kullaniciya gosterilebilir hata."""


_LOCK_HINTS = (
    "kilitled", "permission denied", "another process", "erisemiyor",
    "being used by another", "sharing violation", "kullaniliyor",
)


def _friendly_error(text: str) -> str:
    """Teknik libpff hatasini kullaniciya anlasilir hale getirir."""
    low = text.lower()
    if any(h in low for h in _LOCK_HINTS):
        return (
            "OST dosyasi su anda Outlook tarafindan kullaniliyor (kilitli) "
            "oldugundan dogrudan okunamadi.\n\n"
            "Iki cozumden birini deneyin:\n\n"
            "  1) Outlook'u TAMAMEN kapatin (Gorev Yoneticisi'nde OUTLOOK.EXE "
            "kalmadigindan emin olun), sonra bu islemi yeniden calistirin.\n\n"
            "  2) 'Posta Kutusu -> PST' sekmesini kullanin: Outlook acikken bile "
            "calisir, cunku verinizi Outlook'un kendisi okuyup PST'ye yazar "
            "(dosya kilidi sorunu olmaz). Onerilen yol budur."
        )
    return text


def available_engines() -> Dict[str, bool]:
    """Hangi motorlarin kullanilabilir oldugunu dondurur."""
    return {
        "outlook": engine_outlook.is_available(),
        "libpff": engine_libpff.is_available(),
    }


def list_outlook_stores() -> List["engine_outlook.StoreInfo"]:
    return engine_outlook.list_stores()


# --------------------------------------------------------------------------- #
# OST icerigini kesfetme (arayuzdeki secim agaci icin)
# --------------------------------------------------------------------------- #
def list_ost_tree(ost_path: str) -> List[dict]:
    """OST'nin klasor agacini (govde okumadan) dondurur."""
    if not engine_libpff.is_available():
        raise ConversionError(
            "OST icerigini gostermek icin libpff gereklidir. "
            "('pip install libpff-python')"
        )
    try:
        return engine_libpff.build_tree(ost_path)
    except Exception as exc:
        raise ConversionError(_friendly_error(str(exc))) from exc


def list_ost_folder_messages(ost_path: str, folder_id: str) -> List[dict]:
    """Tek bir klasorun mesaj basliklarini tembel (lazy) listeler."""
    if not engine_libpff.is_available():
        raise ConversionError("libpff gerekli. ('pip install libpff-python')")
    try:
        return engine_libpff.list_folder_messages(ost_path, folder_id)
    except Exception as exc:
        raise ConversionError(_friendly_error(str(exc))) from exc


# --------------------------------------------------------------------------- #
# 1) Bagli posta kutusu (OST onbellegi) -> PST   (tam sadakat, en guvenilir)
# --------------------------------------------------------------------------- #
def mailbox_to_pst(store_id: str, pst_path: str, progress: Progress = None) -> str:
    if not engine_outlook.is_available():
        raise ConversionError(
            "Bu islem icin Windows uzerinde Microsoft Outlook ve pywin32 "
            "gereklidir. ('pip install pywin32')"
        )
    try:
        return engine_outlook.convert_store_to_pst(store_id, pst_path, progress)
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(str(exc)) from exc


# --------------------------------------------------------------------------- #
# 2) Disk uzerindeki .ost dosyasi -> PST
#    libpff ile okunur, Outlook ile gercek PST'ye yazilir.
# --------------------------------------------------------------------------- #
def file_to_pst(ost_path: str, pst_path: str, progress: Progress = None,
                selection: Optional[Dict] = None) -> str:
    if not engine_libpff.is_available():
        raise ConversionError(
            "OST dosyasini okumak icin libpff gereklidir. "
            "('pip install libpff-python')"
        )
    if not engine_outlook.is_available():
        raise ConversionError(
            "Gercek bir PST uretmek icin Outlook + pywin32 gereklidir.\n"
            "Outlook yoksa 'EML / MBOX olarak disa aktar' secenegini kullanin."
        )

    def report(m: str, f: float = -1.0) -> None:
        if progress:
            progress(m, f)

    # --- En guvenilir yol: OST, Outlook'ta TANIMLI bir hesabin dosyasi mi? ---
    # Oyleyse .eml/OpenSharedItem yoluna HIC girmeyiz (bu yol, .eml dosya
    # iliskisi olmayan makinelerde -ozellikle Windows Server- calismaz). Bunun
    # yerine Outlook'tan klasorleri DOGRUDAN kopyalariz (CopyTo): kilit sorunu
    # olmaz, tam sadakatlidir ve .eml gerektirmez.
    matched_store_id = None
    try:
        norm_ost = os.path.normpath(os.path.abspath(ost_path)).lower()
        for st in engine_outlook.list_stores():
            fp = getattr(st, "file_path", "") or ""
            if fp and os.path.normpath(os.path.abspath(fp)).lower() == norm_ost:
                matched_store_id = st.store_id
                break
    except Exception:
        matched_store_id = None

    if matched_store_id:
        report(
            "Bu OST, Outlook'ta tanimli bir hesaba ait. .eml araadimi olmadan, "
            "en guvenilir yontemle (dogrudan kopyalama) aktariliyor. Not: bu "
            "modda tum klasorler kopyalanir (klasor secimi uygulanmaz).",
            -1.0,
        )
        try:
            return engine_outlook.convert_store_to_pst(
                matched_store_id, pst_path, progress
            )
        except ConversionError:
            raise
        except Exception as exc:
            raise ConversionError(_friendly_error(str(exc))) from exc

    # Toplam mesaj sayisini (govde okumadan) hesapla. Ayni zamanda dosyanin
    # acilabilir (kilitli degil) oldugunu DOGRULAR; kilitliyse bos PST uretmeden
    # anlasilir bir hata veririz.
    try:
        total = engine_libpff.count_selected(ost_path, selection)
    except Exception as exc:
        raise ConversionError(_friendly_error(str(exc))) from exc
    if total <= 0:
        raise ConversionError(
            "Secili klasorlerde aktarilacak mesaj bulunamadi.\n"
            "Lutfen mail iceren klasorleri isaretleyin (veya 'Tumunu sec')."
        )

    # Iki asamali, GUVENILIR yontem:
    #   1) Secili tum mesajlar gecici olarak .eml agacina cikarilir.
    #   2) Outlook bunlarin tamamini PST'ye aktarir (dosyalar is bitene kadar
    #      SILINMEZ; aksi halde OpenSharedItem aktarimi tamamlanmadan dosya
    #      silinirse PST bos kalir).
    tmp_dir = tempfile.mkdtemp(prefix="ost2pst_")
    try:
        report("Asama 1/2: OST okunuyor ve gecici olarak cikariliyor...", -1.0)

        def stage1(m: str, f: float) -> None:
            report(m, (f * 0.5) if f >= 0 else f)

        engine_libpff.export_selected_eml(
            ost_path, tmp_dir, selection, stage1, short_names=True
        )

        # TESHIS: 1. asama gercekte kac .eml uretti? Bu tek sayi, bos PST'nin
        # sebebini kesinlestirir (okuma mi, yazma mi).
        n_eml = 0
        for _root, _dirs, _files in os.walk(tmp_dir):
            n_eml += sum(1 for x in _files if x.lower().endswith(".eml"))
        report(f"Asama 1 bitti: {n_eml}/{total} e-posta gecici olarak cikarildi.",
               0.5)
        if n_eml == 0:
            raise ConversionError(
                "OST'nin klasor listesi okundu (%d mesaj gorundu) ANCAK mesaj "
                "GOVDELERI cikarilamadi: 0 e-posta yazildi.\n\n"
                "Bu genellikle su demektir:\n"
                "  - OST su anda Outlook tarafindan aktif/kilitli kullaniliyor, "
                "veya\n"
                "  - bu OST bicimini libpff dogrudan okuyamiyor.\n\n"
                "COZUM: Outlook'ta TANIMLI bu hesap icin 2. sekme "
                "'Posta Kutusu -> PST'yi kullanin. O yontem dosyayi degil, "
                "Outlook'un kendisini okur; kilit/bicim sorunu yasanmaz ve "
                "icerigi tam aktarir." % total
            )

        report("Asama 2/2: Outlook ile PST olusturuluyor...", 0.5)

        def stage2(m: str, f: float) -> None:
            report(m, (0.5 + f * 0.5) if f >= 0 else f)

        result = engine_outlook.import_eml_tree_to_pst(tmp_dir, pst_path, stage2)
        return result
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(_friendly_error(str(exc))) from exc
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# --------------------------------------------------------------------------- #
# 3) Disk uzerindeki .ost dosyasi -> EML agaci veya MBOX  (Outlook gerekmez)
# --------------------------------------------------------------------------- #
def file_to_eml(ost_path: str, out_dir: str, progress: Progress = None,
                selection: Optional[Dict] = None) -> str:
    if not engine_libpff.is_available():
        raise ConversionError("libpff gerekli. ('pip install libpff-python')")
    try:
        return engine_libpff.export_selected_eml(
            ost_path, out_dir, selection, progress
        )
    except Exception as exc:
        raise ConversionError(_friendly_error(str(exc))) from exc


def file_to_mbox(ost_path: str, mbox_path: str, progress: Progress = None,
                 selection: Optional[Dict] = None) -> str:
    if not engine_libpff.is_available():
        raise ConversionError("libpff gerekli. ('pip install libpff-python')")
    try:
        return engine_libpff.export_selected_mbox(
            ost_path, mbox_path, selection, progress
        )
    except Exception as exc:
        raise ConversionError(_friendly_error(str(exc))) from exc


# Geriye donuk uyumluluk icin ince bir sinif sarmalayicisi.
class Conversion:
    """Statik yardimci fonksiyonlara erisim icin basit ad alani."""

    available_engines = staticmethod(available_engines)
    list_outlook_stores = staticmethod(list_outlook_stores)
    list_ost_tree = staticmethod(list_ost_tree)
    list_ost_folder_messages = staticmethod(list_ost_folder_messages)
    mailbox_to_pst = staticmethod(mailbox_to_pst)
    file_to_pst = staticmethod(file_to_pst)
    file_to_eml = staticmethod(file_to_eml)
    file_to_mbox = staticmethod(file_to_mbox)
