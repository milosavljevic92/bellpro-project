Attribute VB_Name = "MdlLocalization"
Option Explicit
Public Sub setLang()
FrmExport.cmdIzvezi.Caption = ""
With FrmInterfejs
    .cmdRefresh.Caption = ""
    .cmdTest.Caption = ""
End With
With FrmMain
    .cmdPrimeni.Caption = ""
    .mnuAbout.Caption = ""
    .mnuinterfejs.Caption = ""
    .mnuizlaz.Caption = ""
    .MnuIzvezi.Caption = ""
    .mnupodesavanja.Caption = ""
    .mnupodprograma.Caption = ""
    .mnupraznici.Caption = ""
    .mnuprogramu.Caption = ""
    .mnuRaspored.Caption = ""
    .mnuraspusti.Caption = ""
    .mnuRegistracija.Caption = ""
    .mnuruucno.Caption = ""
    .mnusvakodnevni.Caption = ""
    .mnutray.Caption = ""
    .MnuUvezi.Caption = ""
    .mnuvannastavne.Caption = ""
    .mnuzastitap.Caption = ""
    .mnuzastitaz.Caption = ""
End With
With FrmNoviRaspored
    .cmdNapravi.Caption = ""
End With
With FrmOtkljucavanje
    .cmdPrijavise.Caption = ""
End With
With FrmPodesavanje
    .chNeZvoni.Caption = ""
    .chNeZvoniNedeljom.Caption = ""
    .chNeZvoniPraznici.Caption = ""
    .chNeZvoniSubotom.Caption = ""
    .chNeZvoniRaspust.Caption = ""
    .chObavestiZvono.Caption = ""
    .chStartMin.Caption = ""
    .chStartWithWin.Caption = ""
    .chVannastavne.Caption = ""
    .chZastitaLozinkom.Caption = ""
    .chZvonoPrekoRazglasa.Caption = ""
End With

With FrmPraznik

End With

With FrmRaspusti

End With

With FrmRegistracija
    .cmdCopy.Caption = ""
    .cmdDemo.Caption = ""
    .cmdRegistruj.Caption = ""
End With

With FrmRucnoZ

End With

With FrmSplash

End With

With FrmSvakodnevni
    .cmdDodaj.Caption = ""
    .cmdNazad.Caption = ""
    .cmdNovi.Caption = ""
    .cmdObrisiVreme.Caption = ""
    .cmdSacuvaj.Caption = ""
End With

End Sub
