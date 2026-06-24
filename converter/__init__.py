"""OST -> PST donusum motorlari.

Bu paket iki bagimsiz motor sunar:

* ``engine_outlook`` : Microsoft Outlook + pywin32 (COM/MAPI) kullanarak
  Outlook'un sorunsuz acabilecegi GERCEK bir PST dosyasi uretir.
  En guvenilir yontemdir; Windows uzerinde Outlook kurulu olmasini gerektirir.

* ``engine_libpff`` : libpff (pypff) kullanarak diskteki bir ``.ost`` dosyasini
  (Outlook profiline bagli olmayan / orphan dosyalar dahil) okur ve disa aktarir.
  Outlook varsa icerik gercek bir PST'ye yazilir, yoksa EML/MBOX olarak kurtarilir.
"""

from .orchestrator import (
    Conversion,
    ConversionError,
    available_engines,
)

__all__ = ["Conversion", "ConversionError", "available_engines"]
