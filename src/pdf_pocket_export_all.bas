Option Explicit


' =========================================================
' PDF EXPORT (ALL SUBFOLDERS, IDW → PDF)
' =========================================================
'
' Поток макроса:
'
' ROOT (папка запуска)
'     ↓
' Берём все подпапки
'     ↓
' Исключаем 1_Архив
'     ↓
' В каждой папке ищем .idw
'     ↓
' Если чертежи есть → считаем папку "пакетом"
'     ↓
' Создаём структуру 2_PDF/<имя папки>
'     ↓
' Открываем каждый .idw
'     ↓
' Настраиваем PDF Translator
'     ↓
' Чистим имя файла
'     ↓
' Экспортируем PDF через SaveCopyAs
'     ↓
' Закрываем документ
'     ↓
' Считаем успехи
'     ↓
' Открываем итоговую папку


Sub Main()

' =========================================================
' 1. НАСТРОЙКИ ПАПОК
' =========================================================

' Базовая папка под PDF (выходная структура)
Dim base_PDF_folder_name As String
base_PDF_folder_name = "2_PDF"

' Архив (игнорируем при обходе)
Dim archive_folder_name As String
archive_folder_name = "1_Архив"


' =========================================================
' 2. ОПРЕДЕЛЕНИЕ КОРНЯ
' =========================================================

' Папка, где лежит текущий документ/запуск макроса
Dim root_dir As String
root_dir = ThisDoc.Path


' =========================================================
' 3. СОЗДАНИЕ PDF-КОРНЯ
' =========================================================

' Формируем путь:
' ROOT → .. → 2_PDF
'
' (поднимаемся на уровень вверх относительно чертежей)

Dim base_PDF_dir As String
base_PDF_dir = System.IO.Path.GetFullPath( _
    System.IO.Path.Combine(root_dir, "..\" & base_PDF_folder_name) _
)

' Если папки нет → создаём
If Not System.IO.Directory.Exists(base_PDF_dir) Then
    System.IO.Directory.CreateDirectory(base_PDF_dir)
End If


' =========================================================
' 4. PDF TRANSLATOR (ENGINE EXPORTA)
' =========================================================

' Берём встроенный Inventor PDF AddIn
Dim pdf_add_in As TranslatorAddIn
pdf_add_in = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

' Контекст трансляции (режим экспорта)
Dim open_context As TranslationContext
open_context = ThisApplication.TransientObjects.CreateTranslationContext
open_context.Type = IOMechanismEnum.kFileBrowseIOMechanism


' =========================================================
' 5. СЧЁТЧИК УСПЕШНЫХ ЭКСПОРТОВ
' =========================================================

Dim total_count As Integer
total_count = 0


' =========================================================
' 6. ПОЛУЧЕНИЕ ПОДПАПОК
' =========================================================

' ROOT
'  ├── subfolder1
'  ├── subfolder2
'  ├── 1_Архив (игнор)
'
Dim sub_directories() As String
sub_directories = System.IO.Directory.GetDirectories(root_dir)

Dim directory_path As String


' =========================================================
' 7. ОБХОД ПАПОК
' =========================================================

For Each directory_path In sub_directories

    Dim directory_name As String
    directory_name = System.IO.Path.GetFileNameWithoutExtension(directory_path)


    ' =====================================================
    ' 7.1 ИСКЛЮЧЕНИЕ АРХИВА
    ' =====================================================

    If directory_name = archive_folder_name Then
        Continue For
    End If


    ' =====================================================
    ' 7.2 ПОИСК IDW
    ' =====================================================

    Dim files() As String
    files = System.IO.Directory.GetFiles(directory_path, "*.idw")

    ' если нет чертежей → это не пакет
    If files.Length = 0 Then
        Continue For
    End If


    ' =====================================================
    ' 7.3 СОЗДАНИЕ PDF ПАПКИ
    ' =====================================================

    Dim curr_pdf_directory As String
    curr_pdf_directory = System.IO.Path.Combine(base_PDF_dir, directory_name)

    If Not System.IO.Directory.Exists(curr_pdf_directory) Then
        System.IO.Directory.CreateDirectory(curr_pdf_directory)
    End If


    ' =====================================================
    ' 8. ОБРАБОТКА ФАЙЛОВ В ПАПКЕ
    ' =====================================================

    Dim curr_file_path As String

    For Each curr_file_path In files

        Dim doc As DrawingDocument
        doc = Nothing


        ' =================================================
        ' 8.1 ОТКРЫТИЕ ЧЕРТЕЖА
        ' =================================================

        Try
            doc = ThisApplication.Documents.Open(curr_file_path, False)


            ' =============================================
            ' 8.2 ИМЯ ФАЙЛА
            ' =============================================

            Dim curr_file_name As String
            curr_file_name = System.IO.Path.GetFileNameWithoutExtension(curr_file_path)


            ' =============================================
            ' 8.3 SANITIZER (защита ФС)
            ' =============================================

            Dim invalid_chars As Char()
            invalid_chars = System.IO.Path.GetInvalidFileNameChars()

            Dim c As Char
            For Each c In invalid_chars
                curr_file_name = curr_file_name.Replace(c, "_"c)
            Next


            ' =============================================
            ' 8.4 ПУТЬ PDF
            ' =============================================

            Dim pdf_path As String
            pdf_path = System.IO.Path.Combine(curr_pdf_directory, curr_file_name & ".pdf")


            ' =============================================
            ' 8.5 ОПЦИИ ЭКСПОРТА
            ' =============================================

            Dim opened_options As NameValueMap
            opened_options = ThisApplication.TransientObjects.CreateNameValueMap

            If pdf_add_in.HasSaveCopyAsOptions(doc, open_context, opened_options) Then
                opened_options.Value("All_Color_AS_Black") = False
                opened_options.Value("Remove_Line_Weights") = False
                opened_options.Value("Vector_Resolution") = 400
                opened_options.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
            End If


            ' =============================================
            ' 8.6 EXPORT ENGINE
            ' =============================================

            Dim opened_data_env As DataMedium
            opened_data_env = ThisApplication.TransientObjects.CreateDataMedium

            opened_data_env.FileName = pdf_path

            pdf_add_in.SaveCopyAs(doc, open_context, opened_options, opened_data_env)


            ' =============================================
            ' 8.7 СЧЁТЧИК
            ' =============================================

            total_count = total_count + 1


        Catch ex As Exception

            MessageBox.Show("Ошибка: " & curr_file_path & vbCrLf & ex.Message)

        Finally

            ' закрываем документ всегда, даже при ошибке
            If Not doc Is Nothing Then doc.Close(True)

        End Try

    Next

Next


' =========================================================
' 9. ФИНАЛ
' =========================================================

MessageBox.Show("Готово. Всего PDF: " & total_count)

System.Diagnostics.Process.Start("explorer.exe", base_PDF_dir)

End Sub
