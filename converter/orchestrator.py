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


def available_engines() -> Dict[str, bool]:
    """Hangi motorlarin kullanilabilir oldugunu dondurur."""
    return {
        "outlook": engine_outlook.is_available(),
        "libpff": engine_libpff.is_available(),
    }


def list_outlook_stores() -> List["engine_outlook.StoreInfo"]:
    return engine_outlook.list_stores()


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
def file_to_pst(ost_path: str, pst_path: str, progress: Progress = None) -> str:
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

    tmp_dir = tempfile.mkdtemp(prefix="ost2pst_")
    try:
        report("Asama 1/2: OST okunuyor ve gecici olarak cikariliyor...", -1.0)

        def stage1(m: str, f: float) -> None:
            report(m, (f * 0.5) if f >= 0 else f)

        engine_libpff.export_eml_tree(ost_path, tmp_dir, stage1)

        report("Asama 2/2: Outlook ile PST olusturuluyor...", 0.5)

        def stage2(m: str, f: float) -> None:
            report(m, (0.5 + f * 0.5) if f >= 0 else f)

        result = engine_outlook.import_eml_tree_to_pst(tmp_dir, pst_path, stage2)
        return result
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(str(exc)) from exc
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# --------------------------------------------------------------------------- #
# 3) Disk uzerindeki .ost dosyasi -> EML agaci veya MBOX  (Outlook gerekmez)
# --------------------------------------------------------------------------- #
def file_to_eml(ost_path: str, out_dir: str, progress: Progress = None) -> str:
    if not engine_libpff.is_available():
        raise ConversionError("libpff gerekli. ('pip install libpff-python')")
    try:
        return engine_libpff.export_eml_tree(ost_path, out_dir, progress)
    except Exception as exc:
        raise ConversionError(str(exc)) from exc


def file_to_mbox(ost_path: str, mbox_path: str, progress: Progress = None) -> str:
    if not engine_libpff.is_available():
        raise ConversionError("libpff gerekli. ('pip install libpff-python')")
    try:
        return engine_libpff.export_mbox(ost_path, mbox_path, progress)
    except Exception as exc:
        raise ConversionError(str(exc)) from exc


# Geriye donuk uyumluluk icin ince bir sinif sarmalayicisi.
class Conversion:
    """Statik yardimci fonksiyonlara erisim icin basit ad alani."""

    available_engines = staticmethod(available_engines)
    list_outlook_stores = staticmethod(list_outlook_stores)
    mailbox_to_pst = staticmethod(mailbox_to_pst)
    file_to_pst = staticmethod(file_to_pst)
    file_to_eml = staticmethod(file_to_eml)
    file_to_mbox = staticmethod(file_to_mbox)
