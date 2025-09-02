; M!Table language file SPANISH
[Printing]
; ***General***
DefPageNr = Pagina %PAGENR%

; ***Status dialog***
StatDlg.Title = Estado Impresora
StatDlg.Text = Imprimiendo p�gina %d...
StatDlg.BtnCancel = &Cancelar

; ***Preview window***
; *Titlebar*
PrevWnd.DefCaption = Previsualizar

; *Toolbar*
PrevWnd.BtnClose = &Cerrar
PrevWnd.BtnPrint = I&mprimir
PrevWnd.BtnPrint.MenuItem.All = Todas
PrevWnd.BtnPrint.MenuItem.CurrPage = Pagina Actual
PrevWnd.BtnScale = &Tama�o
PrevWnd.BtnScale.MenuItem.UserDef = Definido por usuario
PrevWnd.BtnScale.MenuItem.ToPageWid = A ancho de pagina
PrevWnd.BtnFullPage = Pagina Co&mpleta
PrevWnd.BtnFirstPage = &Primera Pagina
PrevWnd.BtnPrevPage = Pagina Pre&via
PrevWnd.BtnNextPage = Pagina &Siguiente
PrevWnd.BtnLastPage = &Ultima Pagina

; *Statusbar*
PrevWnd.StatBar.CurrPage = P�gina %d de %s
PrevWnd.StatBar.CurrScale = Tama�o: %s %%
PrevWnd.StatBar.CreatePage = Creando p�gina %d

; *Messages*
PrevWnd.Msg.Scale.Title = Tama�o
PrevWnd.Msg.Scale.GoToFirstPage.Text = Para cambir el tama�o se debe ir a la primera p�gina. Desea Continuar?
PrevWnd.Msg.Scale.InvValue.Text = Valor Inv�lido.
PrevWnd.Msg.Scale.Input.Text = Ingrese el tama�o en %
PrevWnd.Msg.Scale.Input.BtnOk = Ok
PrevWnd.Msg.Scale.Input.BtnCancel = Cancelar
PrevWnd.Msg.Scale.MaxScaleExceeded.Text = Scaling to page width would exceed the max. scaling of %s %%. Do you like to scale to %s %% instead?
PrevWnd.Msg.Scale.MaxScaleExceeded.Text = Agrandar hasta el ancho de p�gina puede exceder el m�ximo. Tama�o de %s %%. Desea cambiar el tama�o a %s %% de cualquier manera?

[Export]
; ***Status dialog***
StatDlg.Title = Estado Exportaci�n
StatDlg.BtnCancel = &Cancelr
StatDlg.Text.Init = Inicializando
StatDlg.Text.OpenTempl = Abriendo archivo de template...
StatDlg.Text.QueryCols = Buscando columnas a exportar...
StatDlg.Text.ExpColHdrs = Exportando Encabezados de columnas...
StatDlg.Text.GetData = Obteniendo datos...
StatDlg.Text.GetDataProgr = Obteniendo datos... ( %d filas procesadas )
StatDlg.Text.InsData = Insertando Datos...
StatDlg.Text.Format = Formateando...
StatDlg.Text.SaveFile = Guardando Archivo...
StatDlg.Text.CloseWBook = Cerrando Libro...
StatDlg.Text.CloseInst = Cerrando Instancia...

; ***Errors***
Error.CopyToClipboard = Error mientras se copiaba al Clipboard.
Error.CopyToClipboard.Open = Error mientras se copiaba al Clipboard: No se pudo abrir el clipboard.
Error.CopyToClipboard.Alloc = Error mientras se copiaba al Clipboard: No se pudo asignar el objeto en la memoria global.
Error.CopyToClipboard.SetData = Error mientras se copiaba al Clipboard: No se pudieron setear los datos.
Error.GetFontInfo = Error mientras se obtenia informaci�n de Fuentes.
Error.GetObject = Error mientras se obten�a un objeto.
Error.GetStartCell = Error mientras se obetn�a la Celda de comienzo.
Error.InsertSheet = Error mientras se insertaba una Hoja.
Error.SaveClipboard = Error al guardar los datos del clipboard.
Error.SaveClipboard.Open = Error al guardar los datos del clipboard: No se puede abrir el Clipboard.
Error.Unknown = Ocurri� un error desconocido.

; ***Exceptions***
Exception.ErrorCode = C�digo de error:
Exception.OleDispatch.Unknown =  Ocurrio una excepci�n de OLE dispatch.
Exception.Ole.Unknown = Ocurri� un error desconocido de OLE.
Exception.Unknown = Ocurrio una excepci�n desconocida.
Exception.Msg.Title.Excel = Exportaci�n a Excel- Excepci�n
Exception.Msg.Title.OOCalc = Exportacion a Open Office Calc - Excepci�n
Exception.Msg.Title.Default = Exportaci�n - Excepci�n

; ***Open Office Calc***
OOCalc.NewSheet.Name = Hoja