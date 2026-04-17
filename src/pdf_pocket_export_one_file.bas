Option Explicit


' =========================================================
' SINGLE PDF EXPORT (ACTIVE DRAWING)
' WITH RETRY MECHANISM
' =========================================================
'
' Поток макроса:
'
' ACTIVE DOCUMENT
'      ↓
' Проверка DrawingDocument
'      ↓
' Получение part_number (iProperties)
'      ↓
' fallback → имя файла
'      ↓
' sanitizer (очистка имени)
'      ↓
' формирование пути 2_PDF
'      ↓
' настройка PDF translator
'      ↓
' подготовка export options
'      ↓
' попытка SaveCopyAs (x3 retry)
'      ↓
' success / fail output


Sub Main()

' =========================================================
' 1. ПРОВЕРКА ДОКУМЕНТА
' =========================================================

Dim doc As Document = ThisDoc.Document

If Not TypeOf doc Is DrawingDocument Then
    MessageBox.Show("Открытый документ не является чертежом.")
    Exit Sub
End If

Dim drawing_doc As DrawingDocument = doc


' =========================================================
' 2. ПОЛУЧЕНИЕ ИМЕНИ (Part Number)
' =========================================================

Dim part_number As String

Try

    part_number = drawing_doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value

    If part_number = "" Then
        part_number = System.IO.Path.GetFileNameWithoutExtension(doc.DisplayName)
    End If

Catch

    part_number = System.IO.Path.GetFileNameWithoutExtension(doc.DisplayName)

End Try


' =========================================================
' 3. SANITIZER ИМЕНИ ФАЙЛА
' =========================================================

Dim invalid_chars As Char() = System.IO.Path.GetInvalidFileNameChars()

Dim c As Char

For Each c In invalid_chars
    part_number = part_number.Replace(c, "_"c)
Next


' =========================================================
' 4. ФОРМИРОВАНИЕ ПУТИ PDF
' =========================================================

Dim doc_path As String = ThisDoc.Path

Dim pdf_dir As String
pdf_dir = System.IO.Path.GetFullPath( _
    System.IO.Path.Combine(doc_path, "..\..\2_PDF\") _
)

If Not System.IO.Directory.Exists(pdf_dir) Then
    System.IO.Directory.CreateDirectory(pdf_dir)
End If

Dim pdf_path As String
pdf_path = System.IO.Path.Combine(pdf_dir, part_number & ".pdf")


' =========================================================
' 5. PDF TRANSLATOR
' =========================================================

Dim pdf_add_in As TranslatorAddIn
pdf_add_in = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

If pdf_add_in Is Nothing Then
    MessageBox.Show("PDF-плагин не найден.")
    Exit Sub
End If


' =========================================================
' 6. CONTEXT EXPORT
' =========================================================

Dim open_context As TranslationContext
open_context = ThisApplication.TransientObjects.CreateTranslationContext()
open_context.Type = IOMechanismEnum.kFileBrowseIOMechanism


' =========================================================
' 7. EXPORT OPTIONS
' =========================================================

Dim export_options As NameValueMap
export_options = ThisApplication.TransientObjects.CreateNameValueMap()

If Not pdf_add_in.HasSaveCopyAsOptions(doc, open_context, export_options) Then
    MessageBox.Show("Не удалось получить параметры сохранения PDF.")
    Exit Sub
End If

export_options.Value("All_Color_AS_Black") = False
export_options.Value("Remove_Line_Weights") = False
export_options.Value("Vector_Resolution") = 400
export_options.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
export_options.Value("AllSheets") = True


' =========================================================
' 8. DATA MEDIUM (TARGET FILE)
' =========================================================

Dim data_medium As DataMedium
data_medium = ThisApplication.TransientObjects.CreateDataMedium()

data_medium.FileName = pdf_path


' =========================================================
' 9. RETRY MECHANISM
' =========================================================

Dim attempts As Integer = 3
Dim saved As Boolean = False
Dim last_exception As Exception = Nothing

Dim i As Integer

For i = 1 To attempts

    Try

        pdf_add_in.SaveCopyAs(doc, open_context, export_options, data_medium)

        saved = True
        Exit For

    Catch ex As Exception

        last_exception = ex

        System.Threading.Thread.Sleep(1000)

    End Try

Next


' =========================================================
' 10. RESULT
' =========================================================

If saved Then

    MessageBox.Show("PDF успешно сохранён: " & pdf_path)

Else

    MessageBox.Show("Ошибка при сохранении PDF после " & attempts & _
                    " попыток: " & last_exception.Message)

End If

End Sub