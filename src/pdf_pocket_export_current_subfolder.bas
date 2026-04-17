Option Explicit

' =========================================================
' PDF EXPORT (SINGLE FOLDER → 2_PDF/<subfolder>)
' =========================================================
'
' Поток макроса:
'
' ROOT (папка с .idw)
'     ↓
' Получение всех .idw файлов
'     ↓
' Создание структуры 2_PDF/<subfolder_name>
'     ↓
' Открытие каждого чертежа
'     ↓
' Получение Part Number (iProperties)
'     ↓
' fallback → имя файла
'     ↓
' sanitizer (очистка имени)
'     ↓
' формирование PDF пути
'     ↓
' настройка PDF Translator
'     ↓
' SaveCopyAs
'     ↓
' подсчёт успешных экспортов
'     ↓
' открытие папки результата


Sub Main()

' =========================================================
' 1. НАСТРОЙКИ
' =========================================================

Dim base_PDF_folder_name As String = "2_PDF"

' логическая категория экспорта (меняется вручную)
Dim subfolder_name As String = "Обечайка"


' =========================================================
' 2. КОРЕНЬ ПАПКИ
' =========================================================

Dim root_directory As String = ThisDoc.Path


' =========================================================
' 3. СБОРКА PDF СТРУКТУРЫ
' =========================================================

Dim base_PDF_directory As String =
    System.IO.Path.GetFullPath(
        System.IO.Path.Combine(root_directory, "..\..\" & base_PDF_folder_name)
    )

Dim pdf_directory As String =
    System.IO.Path.Combine(base_PDF_directory, subfolder_name)


' =========================================================
' 4. СОЗДАНИЕ ПАПОК
' =========================================================

If Not System.IO.Directory.Exists(base_PDF_directory) Then
    System.IO.Directory.CreateDirectory(base_PDF_directory)
End If

If Not System.IO.Directory.Exists(pdf_directory) Then
    System.IO.Directory.CreateDirectory(pdf_directory)
End If


' =========================================================
' 5. ПОЛУЧЕНИЕ ФАЙЛОВ
' =========================================================

Dim files() As String =
    System.IO.Directory.GetFiles(root_directory, "*.idw")

If files.Length = 0 Then
    MessageBox.Show("Нет .idw файлов")
    Exit Sub
End If


' =========================================================
' 6. PDF TRANSLATOR
' =========================================================

Dim pdf_add_in As TranslatorAddIn =
    ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

Dim opened_context As TranslationContext =
    ThisApplication.TransientObjects.CreateTranslationContext()

opened_context.Type = IOMechanismEnum.kFileBrowseIOMechanism


' =========================================================
' 7. СЧЁТЧИК УСПЕХА
' =========================================================

Dim success_counter As Integer = 0


' =========================================================
' 8. ОБРАБОТКА ФАЙЛОВ
' =========================================================

For Each curr_file_path As String In files

    Dim doc As DrawingDocument = Nothing

    Try

        ' =============================================
        ' 8.1 ОТКРЫТИЕ ЧЕРТЕЖА
        ' =============================================

        doc = ThisApplication.Documents.Open(curr_file_path, False)


        ' =============================================
        ' 8.2 ИМЯ (iProperties → fallback)
        ' =============================================

        Dim curr_part_number As String

        Try
            curr_part_number =
                doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value

            If curr_part_number = "" Then
                curr_part_number = System.IO.Path.GetFileNameWithoutExtension(curr_file_path)
            End If

        Catch
            curr_part_number =
                System.IO.Path.GetFileNameWithoutExtension(curr_file_path)
        End Try


        ' =============================================
        ' 8.3 SANITIZER
        ' =============================================

        Dim invalid_chars As Char() = System.IO.Path.GetInvalidFileNameChars()

        For Each c As Char In invalid_chars
            curr_part_number = curr_part_number.Replace(c, "_"c)
        Next


        ' =============================================
        ' 8.4 PDF PATH
        ' =============================================

        Dim curr_PDF_path As String =
            System.IO.Path.Combine(pdf_directory, curr_part_number & ".pdf")


        ' =============================================
        ' 8.5 OPTIONS
        ' =============================================

        Dim opened_options As NameValueMap =
            ThisApplication.TransientObjects.CreateNameValueMap()

        If pdf_add_in.HasSaveCopyAsOptions(doc, opened_context, opened_options) Then

            opened_options.Value("All_Color_AS_Black") = False
            opened_options.Value("Remove_Line_Weights") = False
            opened_options.Value("Vector_Resolution") = 400
            opened_options.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets

        End If


        ' =============================================
        ' 8.6 EXPORT
        ' =============================================

        Dim opened_data_medium As DataMedium =
            ThisApplication.TransientObjects.CreateDataMedium()

        opened_data_medium.FileName = curr_PDF_path

        pdf_add_in.SaveCopyAs(doc, opened_context, opened_options, opened_data_medium)


        ' =============================================
        ' 8.7 SUCCESS COUNT
        ' =============================================

        success_counter += 1


    Catch ex As Exception

        MessageBox.Show("Ошибка: " & curr_file_path & vbCrLf & ex.Message)

    Finally

        If Not doc Is Nothing Then doc.Close(True)

    End Try

Next


' =========================================================
' 9. RESULT
' =========================================================

MessageBox.Show("Готово: " & success_counter & " / " & files.Length)

System.Diagnostics.Process.Start("explorer.exe", pdf_directory)

End Sub
