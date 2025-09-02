; M!Table language file SPANISH
[Printing]
; ***General***
DefPageNr = Pagina %PAGENR%

; ***Status dialog***
StatDlg.Title = Estado Impresora
StatDlg.Text = Imprimiendo página %d...
StatDlg.BtnCancel = &Cancelar

; ***Preview window***
; *Titlebar*
PrevWnd.DefCaption = Previsualizar

; *Toolbar*
PrevWnd.BtnClose = &Cerrar
PrevWnd.BtnPrint = I&mprimir
PrevWnd.BtnPrint.MenuItem.All = Todas
PrevWnd.BtnPrint.MenuItem.CurrPage = Pagina Actual
PrevWnd.BtnScale = &Tamaño
PrevWnd.BtnScale.MenuItem.UserDef = Definido por usuario
PrevWnd.BtnScale.MenuItem.ToPageWid = A ancho de pagina
PrevWnd.BtnFullPage = Pagina Co&mpleta
PrevWnd.BtnFirstPage = &Primera Pagina
PrevWnd.BtnPrevPage = Pagina Pre&via
PrevWnd.BtnNextPage = Pagina &Siguiente
PrevWnd.BtnLastPage = &Ultima Pagina

; *Statusbar*
PrevWnd.StatBar.CurrPage = Página %d de %s
PrevWnd.StatBar.CurrScale = Tamaño: %s %%
PrevWnd.StatBar.CreatePage = Creando página %d

; *Messages*
PrevWnd.Msg.Scale.Title = Tamaño
PrevWnd.Msg.Scale.GoToFirstPage.Text = Para cambir el tamaño se debe ir a la primera página. Desea Continuar?
PrevWnd.Msg.Scale.InvValue.Text = Valor Inválido.
PrevWnd.Msg.Scale.Input.Text = Ingrese el tamaño en %
PrevWnd.Msg.Scale.Input.BtnOk = Ok
PrevWnd.Msg.Scale.Input.BtnCancel = Cancelar
PrevWnd.Msg.Scale.MaxScaleExceeded.Text = Scaling to page width would exceed the max. scaling of %s %%. Do you like to scale to %s %% instead?
PrevWnd.Msg.Scale.MaxScaleExceeded.Text = Agrandar hasta el ancho de página puede exceder el máximo. Tamaño de %s %%. Desea cambiar el tamaño a %s %% de cualquier manera?

[Export]
; ***Status dialog***
StatDlg.Title = Estado Exportación
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
Error.GetFontInfo = Error mientras se obtenia información de Fuentes.
Error.GetObject = Error mientras se obtenía un objeto.
Error.GetStartCell = Error mientras se obetnía la Celda de comienzo.
Error.InsertSheet = Error mientras se insertaba una Hoja.
Error.SaveClipboard = Error al guardar los datos del clipboard.
Error.SaveClipboard.Open = Error al guardar los datos del clipboard: No se puede abrir el Clipboard.
Error.Unknown = Ocurrió un error desconocido.

; ***Exceptions***
Exception.ErrorCode = Código de error:
Exception.OleDispatch.Unknown =  Ocurrio una excepción de OLE dispatch.
Exception.Ole.Unknown = Ocurrió un error desconocido de OLE.
Exception.Unknown = Ocurrio una excepción desconocida.
Exception.Msg.Title.Excel = Exportación a Excel- Excepción
Exception.Msg.Title.OOCalc = Exportacion a Open Office Calc - Excepción
Exception.Msg.Title.Default = Exportación - Excepción

; ***Open Office Calc***
OOCalc.NewSheet.Name = Hoja