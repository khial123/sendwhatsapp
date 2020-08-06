PARAMETERS NumeroDoTelefoneComDDDSemEspacos, cMsgEnviar, NomeDaImagemBMP, cmodulo, ccomenta

**=enviarwhatsapp(smstelefono, smsmensaje,"", "Inf_clientes","")

IF PARAMETERS()<=2  && no envia BMP
	NomeDaImagemBMP =""
	cmodulo=''
	ccomenta=''
endif
Declare  Integer FindWindow In WIN32API String , String
Declare  Integer SetForegroundWindow In WIN32API Integer
Declare  Integer  ShowWindow  In WIN32API Integer , Integer
Declare Integer ShellExecute In shell32.Dll ;
   INTEGER hndWin, ;
   STRING cAction, ;
   STRING cFileName, ;
   STRING cParams, ;
   STRING cDir, ;
   INTEGER nShowWin

** NumeroDoTelefoneComDDDSemEspacos = 'DD999999999' && Colocar o númeo do celular aqui...
**NumeroDoTelefoneComDDDSemEspacos = '+595984112789' && Colocar o númeo do celular aqui...
** NomeDaImagemBMP = 'D:\TMP\LOGO.BMP' && informar a imagem BMP
!	NomeDaImagemBMP = 'D:\excelsoft\puntoshop\Ayuda\BMPS\LOGOfoxparaguay.BMP' && informar a imagem BMP
**NomeDaImagemBMP ="D:\excelsoft\OtrosSistemas\semei.sueldos\Sueldosemei\fotos\extractogen.pdf"
* cFoneAlvo  = '+54'+NumeroDoTelefoneComDDDSemEspacos  
cFoneAlvo  = NumeroDoTelefoneComDDDSemEspacos
* cMsgEnviar = 'Texto a ser enviado aqui ....'
cMsgEnviar = STRTRAN( cMsgEnviar , CHR(13)+CHR(10), '%0A' )
cMsgEnviar = STRTRAN( cMsgEnviar , ' ', '%20' )

Local lt, lhwnd

comando='whatsapp://send?phone='+Alltrim(cFoneAlvo )+'&text='+cMsgEnviar
=ShellExecute(0, 'open', comando,'', '', 1)

Wait "" Timeout 1
lt    = "WhatsApp"
lhwnd = FindWindow (0, lt)

If lhwnd!= 0 && se estiver com o formulario do WhatsApp ativo
   SetForegroundWindow (lhwnd) && Foca no fomulario do Whatsapp
   ShowWindow (lhwnd, 1)
   ox = Createobject ( "Wscript.Shell" )
	IF FILE(NomeDaImagemBMP)
	   CTRL_C_BMP ( NomeDaImagemBMP  )
	ELSE
		_cliptext = ""
	endif
   Inkey(1)
   ox.sendKeys ( "^{v}" )
   Inkey(1)
   ox.sendKeys ( '{ENTER}' )
   Insert Into whatsapplog;
		(sentlogid,;
		name, ;
		mob, ;
		dttm, ;
		msgtext, ;
		usr, ;
		coment );
		values;
		(myidunico(), ;
		cmodulo,;
		cFoneAlvo, ;
		DATETIME(), ;
		cMsgEnviar, ;
		vs_id,;
		ccomenta)
Else
   MESSAGEBOX ( "Whatsapp no está activo" )
EndIf


Function CTRL_C_BMP( cImageBMP )
lc_Filename=cImageBMP
Local lOk
   Declare Long OpenClipboard in User32 Long nhWnd
   Declare Long EmptyClipboard in User32
   Declare Long CloseClipboard in User32
   Declare Long SetClipboardData in User32 Integer nFormat, Long hMem
   Declare Long LoadImage in User32 ;
       Long hInst, String cFilename, Integer nType, ;
       Integer cxDesired, Integer cyDesired, Integer fuLoad
   lOk = .F.
   hBitmap = LoadImage(0, lc_Filename, 0, 0,0, 0x10)
   If (hBitmap != 0)
      OpenClipboard(0)
      EmptyClipboard()
      SetClipboardData(2, hBitmap)
      CloseClipboard()
      lOk = .T.
   endif
Return lOk
