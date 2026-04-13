/* eslint-disable */
/*
 * Script Name: exportPageRangesToPDF.jsx
 * Description: This InDesign script exports page ranges to separate PDFs and helps with labeling and versioning of the exports. 
 * See _helpText below for details, use case and examples.
 *
 * Author: [Stephan Möbius]
 * Original idea: [Jongware]
 * Multi language L10N-function: Marc Autret, equalizer.js, 2012.
 
 * Changelog and Script on GitHub:
 * https://github.com/MoebiusSt/exportPageRangesToPDF.jsx
 */
 
// Zielsatz-Engine (gibt dem Skript-Kontext eine ID)
//@targetengine "exportPageRangesToPDF"

var SCRIPT_VERSION = "2.09"; // Date: 2026-02-25

// Global debug configuration
var DEBUG_LOG = false;
var logFile = null;

/**
 * Ermittelt das korrekte Pfadtrennzeichen für das aktuelle Betriebssystem
 * @returns {String} Das Pfadtrennzeichen ("\\" für Windows, "/" für Mac)
 */
function getPathSeparator() {
    return ($.os.indexOf("Windows") >= 0) ? "\\" : "/";
}

/**
 * Normalize all path separators in a string to the OS-specific separator
 * @param {String} path
 * @returns {String}
 */
function normalizePathSeparators(path) {
    var sep = getPathSeparator();
    return String(path).replace(/[\/\\]+/g, sep);
}

/**
 * Centralized logging system for the script.
 * Handles all debug output, error reporting, and user alerts.
 */
var Logger = (function() {
    var logFilePath = Folder.desktop + getPathSeparator() + "exportPDF_debug.txt";
    var isLogFileOpen = false;
    
    function openLogFile() {
        if (DEBUG_LOG && !isLogFileOpen) {
            try {
                logFile = new File(logFilePath);
                logFile.open('a'); // Append mode
                isLogFileOpen = true;
                writeToLog("\n--- Debug Log Started " + formatDateForLog(new Date()) + " ---");
                writeToLog("Script Version: " + SCRIPT_VERSION);
                writeToLog("OS: " + $.os);
                return true;
            } catch(e) {
                $.writeln("WARNING: Could not open debug log file: " + e.message);
                return false;
            }
        }
        return false;
    }
    
    function closeLogFile() {
        if (isLogFileOpen && logFile !== null) {
            try {
                writeToLog("--- Debug Log Ended " + formatDateForLog(new Date()) + " ---\n");
                logFile.close();
                isLogFileOpen = false;
                return true;
            } catch(e) {
                $.writeln("WARNING: Could not close debug log file: " + e.message);
                return false;
            }
        }
        return false;
    }
    
    function writeToLog(message) {
        if (isLogFileOpen && logFile !== null) {
            try {
                logFile.writeln(message);
                return true;
            } catch(e) {
                $.writeln("WARNING: Could not write to debug log: " + e.message);
                return false;
            }
        }
        return false;
    }
    
    function formatDateForLog(date) {
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var day = date.getDate();
        var hours = date.getHours();
        var minutes = date.getMinutes();
        var seconds = date.getSeconds();
        
        // Pad single digits with zero
        function pad(num) {
            return (num < 10) ? "0" + num : num.toString();
        }
        
        return year + "-" + pad(month) + "-" + pad(day) + " " + 
               pad(hours) + ":" + pad(minutes) + ":" + pad(seconds);
    }
    
    return {
        init: function() {
            return openLogFile();
        },
        
        close: function() {
            return closeLogFile();
        },
        
        debug: function(message) {
            if (DEBUG_LOG) {
                $.writeln("DEBUG: " + message);
                writeToLog("DEBUG: " + message);
            }
        },
        
        info: function(message) {
            if (DEBUG_LOG) {
                $.writeln("INFO: " + message);
                writeToLog("INFO: " + message);
            }
        },
        
        warn: function(message) {
            if (DEBUG_LOG) {
                $.writeln("WARNING: " + message);
                writeToLog("WARNING: " + message);
            }
        },
        
        error: function(message, error) {
            var errorMsg = message;
            if (error && error.message) {
                errorMsg += ": " + error.message;
            }
            
            $.writeln("ERROR: " + errorMsg);
            if (DEBUG_LOG) {
                writeToLog("ERROR: " + errorMsg);
                if (error && error.stack) {
                    writeToLog("STACK: " + error.stack);
                }
            }
        },
        
        // For critical errors that need to be shown to the user
        userError: function(message, error) {
            var errorMsg = message;
            if (error && error.message) {
                errorMsg += ": " + error.message;
            }
            
            this.error(errorMsg, error);
            alert(errorMsg);
        }
    };
})();

if (typeof _helpText === 'undefined') {
    var _helpText = "This InDesign script is an aid for quickly exporting page ranges to PDFs with labels and versioning. It is particularly usefull in editorial workflows in which you repeatedly export new versions of pages (i.e. of edited articles) for your editors or customers.\n\n" +
    "Features:\n" +
    "- Exports multiple pages or page ranges to individual PDFs. It stores your list for re-use. See EXAMPLE 1.\n" +
    "- You may specify the export-folder and PDF export preset in the settings. Settings are stored with the document.\n" +
    "- You may specify an identifier for each range. (i.e. to name an article). See EXAMPLE 2.\n" +
    "- automatically versions exported PDFs. Optionally overwrite the last PDF-version. See EXAMPLE 3.\n" +
    "- Retrieve a list of page ranges from previous exported PDFs in the export directory. See EXAMPLE 4.\n" +
    "- Retrieve a list of page ranges from the sections and section markers of the active document. See EXAMPLE 5.\n" +
    "- Fetch the page range of the section you are currently viewing (activeSpread) to create a quick export of what you are working on. See EXAMPLE 6.\n" +
    "- You can set a custom Page-Range prefix used in the filename schema (sanitized to be file-system safe).\n" +
    "- Handles odd/even page starts in PDF-view-settings (sets cover sheet YES/NO)\n\n" +
    "__________\n" +
    "File name schema is:\n\n" +
    "[DocumentName]_[prefix]_[PageRange]_[Label-from-brackets]_v[autoVersion].pdf\n\n" +
    "Default prefix: 'Pages'\n\nExample: \"MyDocument_Pages_5-7_Introduction_v1.pdf\"\n\n\n" +
    "__________\n" +
    "Usage: \n\n" +
    "EXAMPLE 1:\n" +
    "Enter a page or page-range per line. Don't use commas. Page Ranges are allowed to overlap. They don't need to be sorted. Enter:\n" +
    "1\n" +
    "5-7\n" +
    "16\n" +
    "5-20\n\n" +
    "EXAMPLE 2:\n" +
    "Specify a label/identifier for a page-range in brackets (optional). Enter:\n\n" +
    "1 (Introduction)\n" +
    "16\n" +
    "17-20(The Beatles)\n\n\n" +
    "EXAMPLE 3:\n" +
    "The version-number will auto increment if a pdf file with the same name exists, i.e. from _v1 to _v2.\n" +
    "In order to overwrite the last version's  PDF of a page range during export, that means if you don't want to increment the version number, add a minus \"-\" to the end of a line. Enter:\n\n" +
    "1(Introduction)-\n" +
    "3\n" +
    "5-7-\n\n\n" +
    "EXAMPLE 4:\n" +
    "Click [-GET FROM DIRECTORY-] to retrieve a list of page ranges from your former exports. The script scans the directory for files matching the naming scheme and generates a list. If it finds \"MyDocument_Pages_1_Introduction_v1.pdf\" and \"MyDocument_Pages_16_v2.pdf\" and \"MyDocument_Pages_21-24_The Doors_v1.pdf\" it would retrieve:\n\n" +
    "1 (Introduction)\n" +
    "16\n" +
    "21-24(The Doors)\n\n\n" +
    "EXAMPLE 5:\n" +
    "Click [-GET FROM SECTIONS-] to retrieve a list of page ranges from the current documents sections and section markers. Use Indesigns pages panels and \"numbering and sections\"-option to specify section markers there in order to manage the labels for the exports. For example: a document has 12 pages and sections starting on 1(always a section on page one), 3, 4 named \"Intro\" and 6 named \"Main topic\" then the script will retrieve:\n\n" +
    "1-2\n" +
    "3\n" +
    "4-5(Intro)\n" +
    "6-12(Main topic)\n\n\n" +
    "EXAMPLE 6:\n" +
    "Click [-SECTIONS FROM THIS SPREAD-] to get page ranges from the active spead of pages you are currently viewing. For example if you are currently viewing the spread [22-23] and this spread is part of a section starting at page 16 named \"Stereolab\" and also a section starting at page 23 up to 26, the script will retrieve:\n\n" +
    "16-22(Stereolab)\n" +
    "23-26\n\n" +
    "Hint:\n" +
    "Click the '-' button next to 'Export PDFs' to automatically append '-' to all lines that already have an exported PDF in the target folder (based on the current filename schema and prefix).\n\n" +
    "__________\n" +
    "Files in Use\n\n" +
    "The script exports PDFs asynchronously for better performance. If some PDF files cannot be exported because they are currently open in another application, the script will continue with the remaining files. After completing all possible exports, it will prompt you to close the blocked files and offer the option to retry those specific exports or cancel the remaining operations.";
}
 
//======================================
// <L10N> :: FRENCH_LOCALE :: GERMAN_LOCALE :: SPANISH_LOCALE :: RUSSIAN_LOCALE :: ARABIC_LOCALE
//======================================
// Version: :: Version: :: Version: :: Versión: :: Версия: :: الإصدار
// Export Page Ranges :: Exporter des plages de pages :: Seitenbereiche exportieren :: Exportar rangos de páginas :: Экспорт диапазонов страниц :: تصدير نطاقات الصفحات
// Help :: Aide :: Hilfe :: Ayuda :: Справка :: مساعدة
// Pages :: Pages :: Seiten :: Páginas :: Страницы :: صفحات
// page range :: plage de pages :: Seitenbereich :: rango de páginas :: диапазон страниц :: نطاق الصفحات
// PDF Folder :: Dossier PDF :: PDF-Ordner :: Carpeta PDF :: Папка PDF :: مجلد PDF
// Page-Range prefix :: Préfixe de plage de pages :: Seitenbereichs-Präfix :: Prefijo de rango de páginas :: Префикс диапазона страниц :: بادئة نطاق الصفحات
// Get from Directory :: Obtenir du répertoire :: Aus Verzeichnis abrufen :: Obtener del directorio :: Получить из каталога :: الحصول من الدليل
// Get from Sections :: Obtenir des sections :: Aus Abschnitten abrufen :: Obtener de secciones :: Получить из разделов :: الحصول من الأقسام
// Sections from this Spread :: Sections de ce planche :: Abschnitte dieses Druckbogens :: Secciones de este pliego :: Разделы этого разворота :: أقسام هذا الانتشار
// Select PDF folder :: Sélectionner le dossier PDF :: PDF-Ordner auswählen :: Seleccionar carpeta PDF :: Выберите папку PDF :: حدد مجلد PDF
// PDF Preset :: Préréglage PDF :: PDF-Vorgabe :: Ajuste preestablecido de PDF :: Предустановка PDF :: الإعداد المسبق لـ PDF
// No files in export folder :: Aucun fichier dans le dossier d'export :: Keine Dateien im Ausgabeverzeichnis :: No hay archivos en la carpeta de exportación :: Нет файлов в папке экспорта :: لا توجد ملفات في مجلد التصدير
// No PDF files in export folder :: Aucun fichier PDF dans le dossier d'export :: Keine PDF-Dateien im Ausgabeverzeichnis :: No hay archivos PDF en la carpeta de exportación :: Нет PDF-файлов в папке экспорта :: لا توجد ملفات PDF في مجلد التصدير
// No PDF exports matching this document found :: Aucun export PDF correspondant à ce document trouvé :: Keine PDF-Exporte passend zu diesem Dokument gefunden :: No se encontraron exportaciones PDF que coincidan con este documento :: Не найдено PDF-экспортов, соответствующих этому документу :: لم يتم العثور على عمليات تصدير PDF مطابقة لهذا المستند
// Different page-range prefix detected: :: Préfixe de plage de pages différent détecté : :: Anderes Seitenbereichs-Präfix erkannt: :: Se detectó un prefijo de rango de páginas diferente: :: Обнаружен другой префикс диапазона страниц: :: تم اكتشاف بادئة نطاق صفحات مختلفة: %1
// Adopt detected prefix into settings? :: Adopter le préfixe détecté dans les paramètres ? :: Soll der erkannte Präfix in die Einstellungen übernommen werden? :: ¿Adoptar el prefijo detectado en la configuración? :: Принять обнаруженный префикс в настройки? :: هل تريد اعتماد البادئة المكتشفة في الإعدادات؟
// No ranges matching current prefix. Found prefixes: :: Aucun intervalle correspondant au préfixe actuel. Préfixes trouvés : :: Keine Seitenbereiche passend zum aktuellen Präfix. Gefundene Präfixe: :: No hay rangos que coincidan con el prefijo actual. Prefijos encontrados: :: Нет диапазонов, соответствующих текущему префиксу. Найденные префиксы: :: لا توجد نطاقات مطابقة للبادئة الحالية. البوادئ الموجودة: %1
// Adopt one of the prefixes into settings? :: Adopter l'un des préfixes dans les paramètres ? :: Soll einer der Präfixe in die Einstellungen übernommen werden? :: ¿Adoptar uno de los prefijos en la configuración? :: Принять один из префиксов в настройки? :: هل تريد اعتماد أحد البوادئ في الإعدادات؟
// OK :: OK :: OK :: OK :: ОК :: موافق
// Error in button click handler :: Erreur dans le gestionnaire de clic de bouton :: Fehler im Button-Klick-Handler :: Error en el controlador de clic del botón :: Ошибка в обработчике нажатия кнопки :: خطأ في معالج النقر على الزر
// Error in getFromDirectory :: Erreur dans getFromDirectory :: Fehler in getFromDirectory :: Error en getFromDirectory :: Ошибка в getFromDirectory :: خطأ في getFromDirectory
// No matching page sections detected from files. Files must match the script's naming scheme :: Aucune section de page correspondante détectée dans les fichiers. Les fichiers doivent correspondre au schéma de nommage du script :: Keine Seitenbereiche aus den Dateien erkannt. Dateien müssen dem Namensschema des Skriptes entsprechen :: No se detectaron secciones de página coincidentes en los archivos. Los archivos deben coincidir con el esquema de nomenclatura del script :: Не обнаружены соответствующие разделы страниц в файлах. Файлы должны соответствовать схеме именования скрипта :: لم يتم اكتشاف أقسام صفحة مطابقة من الملفات. يجب أن تتطابق الملفات مع مخطط التسمية الخاص بالبرنامج النصي
// Incompatible settings were found. Using default settings. :: Des paramètres incompatibles ont été trouvés. Utilisation des paramètres par défaut. :: Inkompatible Einstellungen wurden gefunden. Es werden die Standardeinstellungen verwendet. :: Se encontraron configuraciones incompatibles. Usando configuración predeterminada. :: Обнаружены несовместимые настройки. Используются настройки по умолчанию. :: تم العثور على إعدادات غير متوافقة. استخدام الإعدادات الافتراضية.
// Workfolder from last settings not available. Reverting to default folder. :: Le dossier de travail des derniers paramètres n'est pas disponible. Retour au dossier par défaut. :: Arbeitsordner aus letzten Einstellungen nicht verfügbar. Zurück zum Standardordner. :: La carpeta de trabajo de la última configuración no está disponible. Volviendo a la carpeta predeterminada. :: Рабочая папка из последних настроек недоступна. Возврат к папке по умолчанию. :: مجلد العمل من الإعدادات الأخيرة غير متوفر. العودة إلى المجلد الافتراضي.
// This script requires the InDesign setting 'General > Numbering > Absolute Page Numbering'. It will apply this setting. Do you want to proceed? :: Ce script nécessite le paramètre InDesign 'Général > Numérotation > Numérotation absolue des pages'. Il appliquera ce paramètre. Voulez-vous continuer ? :: Dieses Skript erfordert die InDesign-Einstellung 'Allgemein > Nummerierung > Absolute Seitennummerierung'. Es wird diese Einstellung anwenden. Möchten Sie fortfahren? :: Este script requiere la configuración de InDesign 'General > Numeración > Numeración de página absoluta'. Aplicará esta configuración. ¿Desea continuar? :: Этот скрипт требует настройки InDesign 'Общие > Нумерация > Абсолютная нумерация страниц'. Он применит этот настройку. Хотите продолжить? :: يتطلب هذا البرنامج النصي إعداد InDesign 'عام > الترقيم > ترقيم الصفحات المطلق'. سيتم تطبيق هذا الإعداد. هل تريد المتابعة؟
// Setting required :: Paramètre requis :: Einstellung erforderlich :: Configuración requerida :: Требуется настройка :: الإعداد مطلوب
// optional name :: nom facultatif :: optionaler Name :: nombre opcional :: необязательное имя :: اسم اختياري
// number :: numéro :: Nummer :: número :: номер :: رقم
// An unexpected error occurred in the script. :: Une erreur inattendue s'est produite dans le script. :: Ein unerwarteter Fehler ist im Skript aufgetreten. :: Se produjo un error inesperado en el script. :: В скрипте произошла непредвиденная ошибка. :: حدث خطأ غير متوقع في البرنامج النصي.
// Cancel :: Annuler :: Abbrechen :: Cancelar :: Отмена :: إلغاء
// Settings :: Paramètres :: Einstellungen :: Configuración :: Настройки :: الإعدادات
// Page ranges :: Plages de pages :: Seitenbereiche :: Rangos de página :: Диапазоны страниц :: نطاقات الصفحات
// Export PDFs :: Exporter les PDF :: PDFs exportieren :: Exportar PDFs :: Экспорт PDF :: تصدير ملفات PDF
// Example :: Exemple :: Beispiel :: Ejemplo :: Пример :: مثال
// Invalid input format :: Format d'entrée invalide :: Ungültiges Eingabeformat :: Formato de entrada no válido :: Неверный формат ввода :: تنسيق إدخال غير صالح
// Input must start with a number :: L'entrée doit commencer par un nombre :: Eingabe muss mit einer Zahl beginnen :: La entrada debe comenzar con un número :: Ввод должен начинаться с числа :: يجب أن يبدأ الإدخال برقم
// Invalid page range :: Plage de pages invalide :: Ungültiger Seitenbereich :: Rango de páginas no válido :: Недопустимый диапазон страниц :: نطاق صفحات غير صالح
// Invalid page range format :: Format de plage de pages invalide :: Ungültiges Seitenbereichsformat :: Formato de rango de páginas no válido :: Неверный формат диапазона страниц :: تنسيق نطاق الصفحات غير صالح
// Invalid page numbers :: Numéros de page invalides :: Ungültige Seitenzahlen :: Números de página no válidos :: Недопустимые номера страниц :: أرقام صفحات غير صالحة
// Start page cannot be greater than end page :: La page de début ne peut pas être supérieure à la page de fin :: Startseite kann nicht größer als Endseite sein :: La página de inicio no puede ser mayor que la página final :: Начальная страница не может быть больше конечной :: لا يمكن أن تكون الصفحة الأولى أكبر من الصفحة الأخيرة
// Page range exceeds document's total pages :: La plage de pages dépasse le nombre total de pages du document :: Seitenbereich überschreitet Gesamtseitenzahl des Dokuments :: El rango de páginas excede el total de páginas del documento :: Диапазон страниц превышает общее количество страниц документа :: نطاق الصفحات يتجاوز إجمالي صفحات المستند
// Error in line :: Erreur à la ligne :: Fehler in Zeile :: Error en la línea :: Ошибка в строке :: خطأ في السطر
// %1 of %2 files could not be exported. In most cases, the cause is files that are already open. Please close: :: %1 fichiers sur %2 n'ont pas pu être exportés. Dans la plupart des cas, la cause est des fichiers déjà ouverts: :: %1 von %2 Dateien konnten nicht exportiert werden. In den meisten Fällen sind bereits geöffneten Dateien die Ursache. Schließen Sie: :: %1 de %2 archivos no se pudieron exportar. En la mayoría de los casos, la causa son archivos que ya están abiertos: :: %1 из %2 файлов не удалось экспортировать. В большинстве случаев причина в файлах, которые уже открыты: :: لم يتمكن من تصدير %1 من أصل %2 ملف. في معظم الحالات، السبب هو الملفات المفتوحة بالفعل:
// An error occurred during the export: :: Une erreur s'est produite pendant l'exportation : :: Ein Fehler ist während des Exports aufgetreten: :: Se produjo un error durante la exportación: :: Произошла ошибка во время экспорта: :: حدث خطأ أثناء التصدير:
// No active window detected. Please ensure a document is open and try again. :: Aucune fenêtre active détectée. Veuillez vous assurer qu'un document est ouvert et réessayez. :: Kein aktives Fenster erkannt. Bitte stellen Sie sicher, dass ein Dokument geöffnet ist, und versuchen Sie es erneut. :: No se detectó ninguna ventana activa. Asegúrese de que haya un documento abierto e intente nuevamente. :: Активное окно не обнаружено. Убедитесь, что документ открыт, и повторите попытку. :: لم يتم اكتشاف أي نافذة نشطة. يرجى التأكد من فتح مستند والمحاولة مرة أخرى.
// You are currently viewing a master spread. Master spreads do not contain sections. :: Vous visualisez actuellement un planche type. Les planches types ne contiennent pas de sections. :: Sie betrachten gerade einen Musterdruckbogen. Musterdruckbögen enthalten keine Abschnitte. :: Actualmente está viendo un pliego maestro. Los pliegos maestros no contienen secciones. :: В данный момент вы просматриваете мастер-разворот. Мастер-развороты не содержат разделов. :: أنت تشاهد حاليًا انتشارًا رئيسيًا. الانتشارات الرئيسية لا تحتوي على أقسام.
// An error occurred while retrieving section information: :: Une erreur s'est produite lors de la récupération des informations de section : :: Ein Fehler ist beim Abrufen der Abschnittsinformationen aufgetreten: :: Se produjo un error al recuperar la información de la sección: :: Произошла ошибка при получении информации о разделе: :: حدث خطأ أثناء استرداد معلومات القسم:
// File has no directory path until saved. :: Le fichier n'a pas de chemin de répertoire tant qu'il n'est pas enregistré. :: Datei hat keinen Verzeichnispfad, bis sie gespeichert wird. :: El archivo no tiene ruta de directorio hasta que se guarde. :: У файла нет пути к каталогу, пока он не сохранен. :: ليس للملف مسار دليل حتى يتم حفظه.
// The following files could not be exported: :: Les fichiers suivants n'ont pas pu être exportés : :: Die folgenden Dateien konnten nicht exportiert werden: :: Los siguientes archivos no se pudieron exportar: :: Следующие файлы не удалось экспортировать: :: تعذر تصدير الملفات التالية:
// Please close these files in other programs and click 'Retry' to export again. :: Veuillez fermer ces fichiers dans d'autres programmes et cliquer sur 'Réessayer' pour exporter à nouveau. :: Bitte schließen Sie diese Dateien in anderen Programmen und klicken Sie auf 'Wiederholen', um den Export erneut zu versuchen. :: Por favor, cierre estos archivos en otros programas y haga clic en 'Reintentar' para exportar nuevamente. :: Пожалуйста, закройте эти файлы в других программах и нажмите 'Повторить' для повторного экспорта. :: يرجى إغلاق هذه الملفات في البرامج الأخرى والنقر على 'إعادة المحاولة' للتصدير مرة أخرى.
// Export failed :: Échec de l'exportation :: Export fehlgeschlagen :: Exportación fallida :: Ошибка экспорта :: فشل التصدير
// Retry :: Réessayer :: Wiederholen :: Reintentar :: Повторить :: إعادة المحاولة
// Cancel :: Annuler :: Abbrechen :: Cancelar :: Отмена :: إلغاء
// Incompatible settings were found. Using default settings. :: Des paramètres incompatibles ont été trouvés. Utilisation des paramètres par défaut. :: Inkompatible Einstellungen wurden gefunden. Es werden die Standardeinstellungen verwendet. :: Se encontraron configuraciones incompatibles. Usando configuración predeterminada. :: Обнаружены несовместимые настройки. Используются настройки по умолчанию. :: تم العثور على إعدادات غير متوافقة. استخدام الإعدادات الافتراضية.
// Warning in line :: Avertissement à la ligne :: Warnung in Zeile :: Advertencia en la línea :: Предупреждение в строке :: تحذير في السطر
// Duplicate page range :: Plage de pages en double :: Doppelter Seitenbereich :: Rango de páginas duplicado :: Повторяющийся диапазон страниц :: نطاق صفحات مكرر
// Only the first instance will be exported. :: Seule la première instance sera exportée. :: Nur die erste Instanz wird exportiert. :: Solo se exportará la primera instancia. :: Будет экспортирован только первый экземпляр. :: سيتم تصدير النسخة الأولى فقط.
// Invalid input format (contains control characters) :: Ungültiges Eingabeformat (enthält Steuerzeichen) :: Ungültiges Eingabeformat (enthält Steuerzeichen) :: Formato de entrada no válido (contiene caracteres de control) :: Неверный формат ввода (содержит управляющие символы) :: تنسيق إدخال غير صالح (يحتوي على أحرف تحكم)
// The specified path does not exist and cannot be created: :: Le chemin spécifié n'existe pas et ne peut pas être créé : :: Der angegebene Pfad existiert nicht und kann nicht erstellt werden: :: La ruta especificada no existe y no se puede crear: :: Указанный путь не существует и не может быть создан: :: المسار المحدد غير موجود ولا يمكن إنشاؤه:
// Cannot create directories at the specified path: :: Impossible de créer des répertoires au chemin spécifié : :: Es können keine Ordner am angegebenen Pfad erstellt werden: :: No se pueden crear directorios en la ruta especificada: :: Невозможно создать каталоги по указанному пути: :: لا يمكن إنشاء دلائل في المسار المحدد:
// No write permissions in the specified folder: :: Pas de permissions d'écriture dans le dossier spécifié : :: Keine Schreibrechte im angegebenen Ordner: :: Sin permisos de escritura en la carpeta especificada: :: Нет прав на запись в указанной папке: :: لا توجد أذونات كتابة في المجلد المحدد:
// Path validation error: :: Erreur de validation du chemin : :: Fehler bei der Pfadvalidierung: :: Error de validación de ruta: :: Ошибка проверки пути: :: خطأ في التحقق من صحة المسار:
// </L10N>

var L10N = L10N || (function()
{
    var ln = (function()
    {
        for(var p in Locale) if(Locale[p] == this) return(p);
    }).call(app.locale);
    
    var parseL10N = function(locale_, f)
    {
        var lines = (function()
        { 
            var com = '// ', sep = ' :: ', beg = '<L10N>', end = '</L10N>';
            var l,r = [];
            var uEsc = function(){return String.fromCharCode(Number('0x'+arguments[1]));}
            if( this.open('r') )
            {
                var comSize = com.length;
                while( !this.eof )
                {
                    l = this.readln().replace(/\\u([0-9a-f]{4})/gi, uEsc);
                    if( l.indexOf(com) != 0 ) continue;
                    if( l.indexOf(end) >= 0 ) break;
                    if( l.indexOf(sep) < 0  ) continue;
                    r.push(l.substr(comSize).split(sep));
                }
                this.close();
            }
            while( (l=r.shift()) && l[0] != beg ) {};
            return (l)?[l].concat(r):false;
        }).call(f||File(app.activeScript));
        
        var r=[];
        if (!lines) return r;
        var line = lines[0];
        var locIndex = (function()
        {
            for (var i=1,sz=line.length ; i<sz ; i++)
                if ( line[i] == locale_ ) return i;
            return 0;
        })();
        if (!locIndex) return r;
        
        while( line=lines.shift() )
            if ( typeof line[locIndex] != 'undefined' )
                r[line[0]] = line[locIndex];
        return r;
    }
    
    var tb = parseL10N(ln);
    __ = function(ks){return(tb[ks]||ks);} 
    return {locale: ln};
})();


// Polyfills 
Number.prototype.isEven = function () {
    return (this % 2 == 0) ? true : false;
}
if (!String.prototype.trim) {
    String.prototype.trim = function() {
        return this.replace(/^\s+|\s+$/g, '');
    };
}


/**
 * Removes control characters from a string that could cause issues in section names,
 * including line breaks, zero-width spaces, and other invisible control characters.
 * This function is specifically for cleaning section names and labels, not general text.
 */
function removeControlChars(str) {
    if (str == null) return '';
    
    return String(str)
        .replace(/[\r\n\f\v]/g, '')     // Remove line breaks
        .replace(/\u200B/g, '')         // Remove Zero-Width Space
        .replace(/\u200C/g, '')         // Remove Zero-Width Non-Joiner
        .replace(/\u200D/g, '')         // Remove Zero-Width Joiner
        .replace(/\u2028/g, '')         // Remove Line Separator
        .replace(/\u2029/g, '')         // Remove Paragraph Separator
        .replace(/\uFEFF/g, '')         // Remove Byte Order Mark
        .replace(/[\u0000-\u001F\u007F-\u009F]/g, ''); // Remove other control characters
}

// Function for trimming that can handle both strings and null/undefined
function safeTrim(str) {
    return str == null ? '' : String(str).trim();
}

function fixEncoding(str) {
    if (!str) return '';
    
    var decoded = str.replace(/Ã¤/g, "ä")
                     .replace(/Ã¶/g, "ö")
                     .replace(/Ã¼/g, "ü")
                     .replace(/Ã/g, "ß")
                     .replace(/Ã/g, "Ä")
                     .replace(/Ã/g, "Ö")
                     .replace(/Ã/g, "Ü")
                     .replace(/Ã©/g, "é")
                     .replace(/Ã /g, "à")
                     .replace(/Ã¨/g, "è")
                     .replace(/Ã¹/g, "ù")
                     .replace(/Ãª/g, "ê")
                     .replace(/Ã¢/g, "â")
                     .replace(/Â/g, "");
    
    return removeControlChars(safeTrim(decoded));
}


// Sanitizes a filename segment to be safe on Windows and macOS
function sanitizeFileNameSegment(str) {
    var s = safeTrim(str == null ? '' : String(str));
    // Remove control chars
    s = s.replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
    // Replace invalid filename chars with underscore (Windows + macOS reserved)
    s = s.replace(/[\\\/:\*\?"<>\|]/g, '_');
    // Trim spaces
    s = s.replace(/^\s+|\s+$/g, '');
    // Disallow leading/trailing dots in a segment
    while (s.length && (s.charAt(0) === '.' || s.charAt(s.length - 1) === '.')) {
        if (s.charAt(0) === '.') s = s.substring(1);
        if (s.charAt(s.length - 1) === '.') s = s.substring(0, s.length - 1);
    }
    return s;
}


function checkAndSetAbsolutePageNumbering() {
    if (app.generalPreferences.pageNumbering !== PageNumberingOptions.ABSOLUTE) {
        // Prompt the user with the option to continue or cancel
        var response = confirm(__("This script requires the InDesign setting 'General > Numbering > Absolute Page Numbering'. It will apply this setting. Do you want to proceed?"), true, __("Setting required"));
        
        if (!response) {
            Logger.info("User cancelled absolute page numbering change");
            exit();
        }
        Logger.info("Setting absolute page numbering");
        app.generalPreferences.pageNumbering = PageNumberingOptions.ABSOLUTE;
    }
}


function getFileVersion(path, baseName, identifier) {
    try {
        var files = Folder(path).getFiles(baseName + '*_' + identifier + '_v*.pdf');
        var highestVersion = 0;
        if (files) {
            for (var i = 0; i < files.length; i++) {
                var match = files[i].name.match(/_v(\d+)\.pdf$/);
                if (match) {
                    var version = parseInt(match[1], 10);
                    if (version > highestVersion) {
                        highestVersion = version;
                    }
                }
            }
        }
        Logger.debug("Found highest version " + (highestVersion + 1) + " for " + identifier);
        return highestVersion + 1;
    } catch (e) {
        Logger.error("Error finding file version", e);
        return 1; // Default to version 1 if there's an error
    }
}

function getLocalizedPageIdentifier() {
    var prefix = (typeof settings !== 'undefined' && settings && settings.pageRangePrefix) ? settings.pageRangePrefix : __("Pages");
    prefix = sanitizeFileNameSegment(prefix);
    if (typeof settings !== 'undefined' && settings) settings.pageRangePrefix = prefix; // normalize
    return "_" + prefix + "_";
}


function getFromDirectory(_path) {
    /**
     * Analyzes PDF files in the export directory and generates a list of the most recent page ranges and labels.
     * 
     * This function performs the following steps:
     * 1. Retrieves all PDF files in the specified directory that match the export naming schema.
     * 2. Parses filenames to extract page ranges, labels, and version numbers.
     * 3. For each unique label, keeps track of the file with the highest version number.
     *    In case of a tie in version numbers, it selects the most recently modified file.
     * 4. Sorts the entries based on the first page number of each range.
     * 5. Generates a formatted string with page ranges and labels, matching the required input format.
     * 
     * The function handles URL-encoded characters in filenames (e.g., spaces encoded as %20) by decoding them.
     * It also includes error handling to catch and report any issues during execution.
     * 
     * @returns {string} A formatted string containing page ranges and labels, or an error message if an exception occurs.
     * @throws {Error} If no matching files are found in the directory.
     */
    try {
        Logger.debug("Getting page ranges from directory: " + _path);
        var folder = new Folder(_path);
        var localizedPageIdentifier = getLocalizedPageIdentifier();
        var docName = d.name.replace(/\..+$/, '');
        var filePattern = escapeRegExp(docName + localizedPageIdentifier) + ".+_v\\d+\\.pdf$";

        // Comfort checks
        var allEntries = folder.getFiles();
        if (!allEntries || allEntries.length === 0) {
            Logger.warn("Export dir contains no files");
            return __("No files in export folder");
        }

        var pdfEntries = [];
        var i;
        for (i = 0; i < allEntries.length; i++) {
            var e = allEntries[i];
            if (e instanceof File && /\.pdf$/i.test(String(e.name))) {
                pdfEntries[pdfEntries.length] = e;
            }
        }
        if (pdfEntries.length === 0) {
            Logger.warn("Export dir contains no PDF files");
            return __("No PDF files in export folder");
        }

        var startsWithDoc = [];
        var startsWithDocRe = new RegExp("^" + escapeRegExp(docName + "_"), "i");
        for (i = 0; i < pdfEntries.length; i++) {
            if (startsWithDocRe.test(decodeURI(pdfEntries[i].name))) {
                startsWithDoc[startsWithDoc.length] = pdfEntries[i];
            }
        }
        if (startsWithDoc.length === 0) {
            Logger.warn("No PDFs matching this document base name found");
            return __("No PDF exports matching this document found");
        }

        // Matching PDFs for the current prefix
        var files = folder.getFiles(function(file) {
            return decodeURI(file.name).match(new RegExp(filePattern, "i"));
        });
        
        if (files.length === 0) {
            // Try to detect other page-range prefixes from existing PDFs of this doc
            var prefixRegex = new RegExp("^" + escapeRegExp(docName + "_") + "([^_]+)_(\\d+-?\\d*)(?:_.*)?_v\\d+\\.pdf$", "i");
            var foundPrefixesMap = {};
            var foundPrefixes = [];
            for (i = 0; i < startsWithDoc.length; i++) {
                var nm = decodeURI(startsWithDoc[i].name);
                var m = nm.match(prefixRegex);
                if (m && m[1]) {
                    var pfx = safeTrim(m[1]);
                    if (!foundPrefixesMap[pfx]) {
                        foundPrefixesMap[pfx] = true;
                        foundPrefixes[foundPrefixes.length] = pfx;
                    }
                }
            }
            // Remove current prefix if present
            var currentPrefix = localizedPageIdentifier.replace(/^_/, '').replace(/_$/, '');
            var remainingPrefixes = [];
            for (i = 0; i < foundPrefixes.length; i++) {
                if (foundPrefixes[i] !== currentPrefix) remainingPrefixes[remainingPrefixes.length] = foundPrefixes[i];
            }

            if (remainingPrefixes.length === 1) {
                // Offer to adopt the single detected prefix
                var detected = remainingPrefixes[0];
                // Bring InDesign to the foreground before showing the modal dialog.
                // Without this, the dialog can appear hidden behind other windows on Windows,
                // blocking all InDesign interaction with no way to recover.
                app.activate();
                
                var adoptSingleDlg = new Window("dialog", __("Page-Range prefix"));
                adoptSingleDlg.orientation = "column";
                adoptSingleDlg.alignChildren = ["fill","top"];
                adoptSingleDlg.margins = 16;
                var msg = adoptSingleDlg.add("statictext", undefined, __("Different page-range prefix detected:") + " " + detected, {multiline:true});
                msg.preferredSize.width = 380;
                var q = adoptSingleDlg.add("statictext", undefined, __("Adopt detected prefix into settings?"));
                var btns = adoptSingleDlg.add("group");
                btns.orientation = "row";
                btns.alignment = ["center","center"];
                var okB = btns.add("button", undefined, __("OK"));
                var cancelB = btns.add("button", undefined, __("Cancel"));
                okB.onClick = function(){ adoptSingleDlg.close(1); };
                cancelB.onClick = function(){ adoptSingleDlg.close(0); };
                var res = adoptSingleDlg.show();
                if (res == 1) {
                    settings.pageRangePrefix = detected;
                    try { if (w && w.tabs && w.tabs.settingsPanel && w.tabs.settingsPanel.prefixGroup) { w.tabs.settingsPanel.prefixGroup.prefixInput.text = detected; } } catch(e){}
                    // Re-run with new prefix
                    return getFromDirectory(_path);
                }
                // fall through to generic message below if user cancels
            } else if (remainingPrefixes.length > 1) {
                // Let user pick one of the detected prefixes
                var chooseDlg = new Window("dialog", __("Page-Range prefix"));
                chooseDlg.orientation = "column";
                chooseDlg.alignChildren = ["fill","top"];
                chooseDlg.margins = 16;
                var msg2 = chooseDlg.add("statictext", undefined, __("No ranges matching current prefix. Found prefixes:") + " " + remainingPrefixes.join(", "), {multiline:true});
                msg2.preferredSize.width = 420;
                var prompt = chooseDlg.add("statictext", undefined, __("Adopt one of the prefixes into settings?"));
                var dd = chooseDlg.add("dropdownlist");
                for (i = 0; i < remainingPrefixes.length; i++) { dd.add('item', remainingPrefixes[i]); }
                dd.selection = 0;
                var btns2 = chooseDlg.add("group");
                btns2.orientation = "row";
                btns2.alignment = ["center","center"];
                var ok2 = btns2.add("button", undefined, __("OK"));
                var cancel2 = btns2.add("button", undefined, __("Cancel"));
                ok2.onClick = function(){ chooseDlg.close(1); };
                cancel2.onClick = function(){ chooseDlg.close(0); };
                var res2 = chooseDlg.show();
                if (res2 == 1 && dd.selection) {
                    var selected = String(dd.selection.text);
                    settings.pageRangePrefix = selected;
                    try { if (w && w.tabs && w.tabs.settingsPanel && w.tabs.settingsPanel.prefixGroup) { w.tabs.settingsPanel.prefixGroup.prefixInput.text = selected; } } catch(e){}
                    return getFromDirectory(_path);
                }
            }

            Logger.warn("No matching files found for current prefix in directory: " + _path);
            return __("No matching page sections detected from files. Files must match the script's naming scheme") + ": \r\n\r\n" +
                   __("Example") + ":\r\n" +
                   "\"" + d.name.replace(/\..+$/, '') + localizedPageIdentifier + "11-12_Advertorial_v02.pdf" + "\""; 
        }

        Logger.info("Found " + files.length + " matching files");
        var pageRanges = {};
        for (var i = 0; i < files.length; i++) {
            var file = files[i];
            var fileName = file.name;
            var regex = new RegExp("^.+" + escapeRegExp(localizedPageIdentifier) + "(\\d+-?\\d*)_?(.*)_v(\\d+)\\.pdf$", "i");
            var match = fileName.match(regex);

            if (match) {
                var pageRange = safeTrim(match[1]);
                var label = fixEncoding(decodeURIComponentSafe(safeTrim(match[2])));
                
                // If label is empty, keep it empty; use pageRange as a key but don't display it as a label
                var identifier = label || pageRange;
                
                var version = parseInt(match[3], 10);
                
                Logger.debug("Found file: " + fileName + " - Page range: " + pageRange + ", Label: " + label + ", Version: " + version);
                
                // Keep the highest version (or most recent if versions are equal)
                if (!pageRanges[identifier] || version > pageRanges[identifier].version || 
                    (version === pageRanges[identifier].version && file.modified > pageRanges[identifier].modified)) {
                    pageRanges[identifier] = {
                        pageRange: pageRange,
                        label: label, // Store the original label separately
                        version: version,
                        modified: file.modified
                    };
                }
            }
        }
        
        // Convert to array and sort by first page number
        var sortedEntries = [];
        for (var identifier in pageRanges) {
            sortedEntries.push({
                label: pageRanges[identifier].label,
                pageRange: pageRanges[identifier].pageRange,
                firstPage: parseInt(pageRanges[identifier].pageRange.split('-')[0], 10)
            });
        }
        
        sortedEntries.sort(function(a, b) {
            return a.firstPage - b.firstPage;
        });
        
        // Generate the formatted result string
        var result = '';
        for (var i = 0; i < sortedEntries.length; i++) {
            if (sortedEntries[i].label) {
                result += sortedEntries[i].pageRange + '(' + sortedEntries[i].label + ')\n';
            } else {
                result += sortedEntries[i].pageRange + '\n';
            }
        }
        
        Logger.info("Retrieved " + sortedEntries.length + " page ranges from directory");
        return safeTrim(result);
    } catch (error) {
        Logger.error("Error in getFromDirectory", error);
        alert(__("Error in getFromDirectory") + ": " + error.message);
        return __("An unexpected error occurred in the script.");
    }
}

function getSectionPages() {
    try {
        Logger.debug("Getting page ranges from document sections");
        var sections = d.sections;
        var sectionInfo = [];
        
        for (var i = 0; i < sections.length; i++) {
            var section = sections[i];
            var start = section.pageStart.documentOffset + 1;
            var end = start + section.length - 1;
            
            var sectionName = createSectionName(section);
            Logger.debug("Found section: " + start + "-" + end + (sectionName ? " (" + sectionName + ")" : ""));
            
            sectionInfo.push({
                start: start,
                end: end,
                name: sectionName
            });
        }
        
        // Sort the sections according to the start value
        sectionInfo.sort(function(a, b) {
            return a.start - b.start;
        });
        
        // Formatting the sorted sections
        var result = [];
        for (var j = 0; j < sectionInfo.length; j++) {
            var info = sectionInfo[j];
            result.push(formatSectionRange(info.start, info.end, info.name));
        }
        
        Logger.info("Retrieved " + sectionInfo.length + " sections from document");
        return result.join("\n");
    } catch (error) {
        Logger.error("Error getting section pages", error);
        return "";
    }
}

function createSectionName(section) {
    var parts = [];
    if (section.includeSectionPrefix && section.sectionPrefix !== "") {
        parts.push(removeControlChars(safeTrim(section.sectionPrefix)));
    }
    if (section.marker !== "") {
        parts.push(removeControlChars(safeTrim(section.marker)));
    }
    return parts.join("-");
}

function formatSectionRange(start, end, sectionName) {
    var range = start === end ? start.toString() : start + "-" + end;
    return sectionName ? range + "(" + sectionName + ")" : range;
}

function getActiveSection() {
    try {
        Logger.debug("Getting sections from active spread");
        var activeWindow = app.activeWindow;

        if (!activeWindow) {
            Logger.warn("No active window detected");
            alert(__("No active window detected. Please ensure a document is open and try again."));
            return "";
        }
        var activeSpread = activeWindow.activeSpread;
        var activePage = activeWindow.activePage;
        // Check whether we are on a MasterPage
        if ((activeSpread && activeSpread.parent instanceof MasterSpread) ||
            (activePage && activePage.parent instanceof MasterSpread)) {
            Logger.warn("User is viewing master spread");
            alert(__("You are currently viewing a master spread. Master spreads do not contain sections."));
            return "";
        }
        
        var spreadSections = [];
        var seenSections = {};

        try {
            var sections = activeSpread.pages.everyItem().appliedSection;
            for (var i = 0; i < sections.length; i++) {
                var section = sections[i];
                if (!seenSections[section.id]) {
                    var cleanSectionName = createSectionName(section);
                    Logger.debug("Found section in spread: " + 
                                (section.pageStart.documentOffset + 1) + "-" + 
                                (section.pageStart.documentOffset + section.length) + 
                                (cleanSectionName ? " (" + cleanSectionName + ")" : ""));
                    
                    spreadSections.push(formatSectionRange(
                        section.pageStart.documentOffset + 1, 
                        section.pageStart.documentOffset + section.length, 
                        cleanSectionName
                    ));
                    seenSections[section.id] = true;
                }
            }
            Logger.info("Retrieved " + spreadSections.length + " sections from active spread");
        } catch (e) {
            Logger.error("Error retrieving section information", e);
            alert(__("An error occurred while retrieving section information: ") + e.message);
            return "";
        }
        return spreadSections.join("\n");
    } catch (error) {
        Logger.error("Error in getActiveSection", error);
        return "";
    }
}



/* ################## MAIN ################## */

// Initialize the logger at the beginning of the script
Logger.init();
Logger.info("Script starting - Version " + SCRIPT_VERSION);

var d = app.activeDocument;

var settings = loadSettings(d) || {
    /* DEFAULT SETTINGS ... Note: if we change the schema of the settings then remember to adjust the loadSettings validation! */
    pageRanges: "",
    folder: "/",
    exportPreset: "[PDF/X-4:2008]",
    pageRangePrefix: __("Pages"),
    windowBounds: undefined
};

checkAndSetAbsolutePageNumbering();

var _lastRange = settings.pageRanges;
var _folder = settings.folder;
var _exportPreset = settings.exportPreset;

var globalValidatedInputs = [];
var _backupViewPDF = app.pdfExportPreferences.viewPDF;
app.pdfExportPreferences.viewPDF = false;



// *** THE DIALOG ***

// Define the dialog as a complete string - this is the only way to get ScriptUI render the ui elements and sizes correctly.
var dialogString = 
"dialog { \
    text: '" + __("Export Page Ranges") + " – " + __("Version:") + " " + SCRIPT_VERSION + "', \
    orientation:'column', \
    margins:0, \
    spacing: 0,\
    preferredSize: [" + (settings.windowBounds ? settings.windowBounds.width : 240) + ", " + (settings.windowBounds ? settings.windowBounds.height : 280) + "], \
    " + (settings.windowBounds ? "location: [" + settings.windowBounds.x + ", " + settings.windowBounds.y + "], " : "") + "\
    alignChildren: ['fill', 'fill'], \
    properties: {closeButton: true, maximizeButton: false, minimizeButton: false, resizeable: true}, \
    mainGroup : Group {\
        margins:[8,10,8,0], \
        alignment: ['fill', 'top'], \
        buttonGroup: Group { \
            alignment: ['right', 'top'], \
            orientation: 'row', \
            margins: 0, \
            spacing: 5, \
            settingsButton: Button { \
                text: '" + __("Settings") + "', \
                properties: {name: 'Settings'}, \
                margins:0 \
            }, \
            helpButton: Button { \
                text: '?', \
                properties: {name: 'Help'}, \
                margins:0, \
                preferredSize: [30, -1] \
            } \
        } \
    }, \
    tabs : Group { \
        alignChildren:['fill','fill'], \
        orientation:'stack', spacing:0, margins:0, \
        minimumSize: [245, 260], \
        mainPanel: Panel { \
            text: '" + __("Page ranges") + "', \
            spacing: 15, \
            margins:[8,10,8,10], \
            editGroup: Group { \
                orientation: 'column', \
                alignment: ['fill', 'fill'], \
                pageRangeEntries: EditText { \
                    properties: {multiline: true, scrollable: true}, \
                    alignment: ['fill', 'fill'], \
                    preferredSize: [160,160],\
                } \
            }, \
            buttonGroup: Group { \
                orientation: 'column', \
                alignChildren: ['center', 'top'], \
                alignment: ['fill', 'bottom'], \
                getFromDirButton: Button { text: '" + __("Get from Directory") + "' }, \
                getFromSectionsButton: Button { text: '" + __("Get from Sections") + "' }, \
                getThisSectionButton: Button { text: '" + __("Sections from this Spread") + "' },\
                actionGroup: Group { \
                    orientation: 'row', \
                    alignChildren: ['center', 'center'], \
                    spacing: 10, \
                    minusButton: Button { text: '-', preferredSize: [30, -1] }, \
                    okButton: Button { text: '" + __("Export PDFs") + "', properties: {name: 'ok'} }, \
                    cancelButton: Button { text: '" + __("Cancel") + "', properties: {name: 'cancel'} } \
                } \
            } \
        }, \
        settingsPanel: Group { \
            margins:[8,0,8,0], \
            orientation: 'column', \
            visible: false, \
            alignChildren: ['fill', 'top'], \
            folderGroup: Group { \
                orientation: 'column', \
                margins:[0,0,0,0], \
                alignChildren: ['left', 'center'], \
                folderLabel: StaticText { text: '" + __("PDF Folder") + "' }, margins:[0,0,0,0], \
                folderInput: EditText { alignment: ['fill', 'center'] }, \
                folderButton: Button { text: '" + __("Select PDF folder") + "' } \
            }, \
            prefixGroup: Group { \
                orientation: 'column', \
                margins:[0,10,0,0], \
                alignChildren: ['left', 'center'], \
                prefixLabel: StaticText { text: '" + __("Page-Range prefix") + "' }, \
                prefixInput: EditText { alignment: ['fill', 'center'] } \
            }, \
            presetGroup: Group { \
                orientation: 'column', \
                preferredSize: [160,20],\
                alignChildren: ['left', 'center'], \
                presetLabel: StaticText { text: '" + __("PDF Preset") + ":' }, margins:[0,15,0,0], \
                presetDropdown: DropDownList { preferredSize: [160,20], alignment: ['fill', 'center'] } \
            } \
        } \
    } \
}";

// Create the dialog
var w = new Window(dialogString);

// Set initial size
w.tabs.minimumSize = [240, 250];

// Populate PDF preset dropdown
var presets = app.pdfExportPresets.everyItem().name;
for (var i = 0; i < presets.length; i++) {
    w.tabs.settingsPanel.presetGroup.presetDropdown.add('item', presets[i]);
    if (presets[i] === _exportPreset) {
        w.tabs.settingsPanel.presetGroup.presetDropdown.selection = i;
    }
}

// Set initial values
w.tabs.settingsPanel.folderGroup.folderInput.text = _folder;
w.tabs.mainPanel.editGroup.pageRangeEntries.text = _lastRange;
w.tabs.settingsPanel.prefixGroup.prefixInput.text = settings.pageRangePrefix || __("Pages");

// live sanitize prefix input
w.tabs.settingsPanel.prefixGroup.prefixInput.onChanging = function() {
    var s = sanitizeFileNameSegment(this.text);
    if (!s) s = __("Pages");
    if (this.text !== s) this.text = s;
};


// Resize event handler
w.onResizing = w.onResize = function() {
    this.layout.resize();
}

// onClose event to save the window position and size
w.onClose = function() {
    var b = this.bounds;
    // Convert ScriptUI bounds [left, top, right, bottom] to an object
    settings.windowBounds = { x: b[0], y: b[1], width: (b[2] - b[0]), height: (b[3] - b[1]) };
};

// Define behavior for the SETTINGS button  
w.mainGroup.buttonGroup.settingsButton.onClick = function() {
    try {
        if (w.tabs.settingsPanel.visible == false) {
            w.tabs.settingsPanel.visible = true;
            w.tabs.mainPanel.visible = false;
            // Disable OK while settings are visible to prevent accidental submit via Enter
            try { w.tabs.mainPanel.buttonGroup.actionGroup.okButton.enabled = false; } catch(_e) {}
        }
        else { 
            // validate path before exiting settings-page
            if (validateFolderPath()) {
                settings.folder = w.tabs.settingsPanel.folderGroup.folderInput.text;
                var pr = sanitizeFileNameSegment(w.tabs.settingsPanel.prefixGroup.prefixInput.text);
                settings.pageRangePrefix = pr || __("Pages");
                w.tabs.settingsPanel.visible = false;
                w.tabs.mainPanel.visible = true;
                // Re-enable OK when returning to main panel
                try { w.tabs.mainPanel.buttonGroup.actionGroup.okButton.enabled = true; } catch(_e) {}
            }
        }
    } catch (e) {
        Logger.userError(__("Error in button click handler"), e);
    }
}; 

w.mainGroup.buttonGroup.helpButton.onClick = function() {
    var helpDialog = new Window("dialog", __("Help"), undefined, {resizeable: true});
    helpDialog.preferredSize = [500, 600];
    helpDialog.minimumSize = [300, 400];
    helpDialog.maximumSize = [500, 600];
    helpDialog.alignChildren = ["fill", "fill"];
    helpDialog.spacing = 5;
    
    var helpText = helpDialog.add("edittext", undefined, _helpText, {
        multiline: true, 
        readonly: true,
        scrolling: true
    });
    helpText.alignment = ["fill", "fill"];
    
    helpDialog.onResize = function() {
        this.layout.resize();
    }
    
    helpDialog.center(w);
    helpDialog.show();
    
    // Dialog danach aufräumen
    if (helpDialog && helpDialog.isValid) {
        helpText = null;
        helpDialog.destroy();
        helpDialog = null;
    }
}

function withButtonWorkaround(button, fn) {// higher order function for a workaround to unhighlight lit buttons
    return function() {
        button.active = true;
        try {
            fn.apply(this, arguments);
        } catch (error) {
            Logger.userError(__("Error in button click handler"), error);
        }
        button.active = false;
    };
}


// Folder button click handler
w.tabs.settingsPanel.folderGroup.folderButton.onClick = withButtonWorkaround(
    w.tabs.settingsPanel.folderGroup.folderButton,
    function() {
        if (d.isValid && d.saved && d.fullName !== null) {
            //preparing startpath for the following selectDlg
            var defaultFolder = getFormattedPath(w.tabs.settingsPanel.folderGroup.folderInput.text, d.fullName.path);
            if (!defaultFolder.exists) {
                defaultFolder = Folder(d.fullName.path); 
            }
            //now let user select a folder
            var selectedFolder = defaultFolder.selectDlg(__("Select PDF folder"));
            if (selectedFolder) {
                var selectedFs = String(selectedFolder.fsName);
                var baseFs = String(d.filePath.fsName);
                var sep = getPathSeparator();
                var resultPath = "";

                if (selectedFs.indexOf(baseFs) === 0) {
                    resultPath = selectedFs.substring(baseFs.length);
                    if (resultPath === "") resultPath = sep; // same folder
                } else {
                    resultPath = selectedFs;
                }

                resultPath = normalizePathSeparators(resultPath);

                // Ensure absolute POSIX style on macOS
                if ($.os.indexOf("Windows") < 0) {
                    if (selectedFs.indexOf(baseFs) !== 0 && resultPath.charAt(0) !== "/") {
                        resultPath = "/" + resultPath.replace(/^\/+/, '');
                    }
                }

                if (resultPath.charAt(resultPath.length - 1) !== sep) {
                    resultPath += sep;
                }

                w.tabs.settingsPanel.folderGroup.folderInput.text = settings.folder = resultPath;
            }
        }
        else {
            throw new Error(__("File has no directory path until saved."));
        }
    }
);

// Get from Directory button click handler
w.tabs.mainPanel.buttonGroup.getFromDirButton.onClick = withButtonWorkaround(
    w.tabs.mainPanel.buttonGroup.getFromDirButton,
    function() {
        if (d.isValid && d.saved && d.fullName !== null) {
            var _path = getFormattedPath(w.tabs.settingsPanel.folderGroup.folderInput.text, d.fullName.path);
            var result = getFromDirectory(_path);
            w.tabs.mainPanel.editGroup.pageRangeEntries.text = result;
        }
        else throw new Error(__("File has no directory path until saved."));
    }
);

// Get from Sections button click handler
w.tabs.mainPanel.buttonGroup.getFromSectionsButton.onClick = withButtonWorkaround(
    w.tabs.mainPanel.buttonGroup.getFromSectionsButton,
    function() {
        var result = getSectionPages();
        w.tabs.mainPanel.editGroup.pageRangeEntries.text = result;
    }
);

w.tabs.mainPanel.buttonGroup.getThisSectionButton.onClick = withButtonWorkaround(
    w.tabs.mainPanel.buttonGroup.getThisSectionButton, 
    function() {
        var result = getActiveSection();
        if (result) {
            w.tabs.mainPanel.editGroup.pageRangeEntries.text = result;
        }
    }
);

// New single [-] button handler near Export PDFs: append minus to lines that already exist in export dir
w.tabs.mainPanel.buttonGroup.actionGroup.minusButton.onClick = withButtonWorkaround(
    w.tabs.mainPanel.buttonGroup.actionGroup.minusButton,
    function() { appendMinusForExistingExports(); }
);

// Helper: check if export exists for a given line (pageRange/label)
function exportExists(_path, pageRange, label) {
    try {
        var folder = new Folder(_path);
        var localizedPageIdentifier = getLocalizedPageIdentifier();
        var baseName = String(d.name).replace(/\..+$/, '');
        var safeLabel = label ? sanitizeFileNameSegment(label) : '';
        var fileBase = baseName + localizedPageIdentifier + pageRange + (safeLabel ? "_" + safeLabel : "");
        var regex = new RegExp("^" + escapeRegExp(fileBase) + "_v\\d+\\.pdf$", "i");
        var files = folder.getFiles(function(file) { return regex.test(decodeURI(file.name)); });
        return files && files.length > 0;
    } catch (e) {
        Logger.error("Error checking export existence", e);
        return false;
    }
}

function appendMinusForExistingExports() {
    if (!(d.isValid && d.saved && d.fullName !== null)) return;
    var _path = getFormattedPath(w.tabs.settingsPanel.folderGroup.folderInput.text, d.fullName.path);
    var lines = String(w.tabs.mainPanel.editGroup.pageRangeEntries.text).split(/\n/);
    var out = [];
    for (var i = 0; i < lines.length; i++) {
        var line = safeTrim(lines[i]);
        if (!line) { out[out.length] = line; continue; }
        var m = line.match(/^([^\(]+?)(?:\(([^\)]*)\))?(\-)?$/);
        if (m) {
            var pr = safeTrim(m[1]);
            var label = m[2] ? safeTrim(m[2]) : '';
            var hasMinus = m[3] === '-';
            if (exportExists(_path, pr, label)) {
                if (!hasMinus) line = pr + (label ? '(' + label + ')' : '') + '-';
            }
        }
        out[out.length] = line;
    }
    w.tabs.mainPanel.editGroup.pageRangeEntries.text = out.join('\n');
}

w.tabs.mainPanel.buttonGroup.actionGroup.okButton.onClick = function() {
    if (validateInputs()) {
        w.close(1); // Close the window with return value 1 if the validation was successful
    }
    // If validateInputs returns false, the window remains open
}

// *** END DIALOG **


// Input validation function
function validateAndFormatInput(input) {
    var trimmedInput = safeTrim(input);
    var totalPages = d.pages.length;
    
    var cleanedInput = removeControlChars(trimmedInput);
    if (cleanedInput !== trimmedInput) {
        trimmedInput = cleanedInput;
    }
    
    // Check whether the input starts with a number
    if (!/^\d/.test(trimmedInput)) {
        throw new Error(__("Input must start with a number"));
    }

    // Split the input into page area and label parts
    var match = trimmedInput.match(/^(\d+(?:\s*-\s*\d+)?)\s*(?:\(([^)]*)\))?\s*(-)?$/);
    if (!match) {
        throw new Error(__("Invalid input format") + 
                         (trimmedInput !== input ? " (" + __("contains control characters") + ")" : ""));
    }
    
    var pageRange = match[1] ? match[1].replace(/\s+/g, '') : '';
    var label = match[2] ? safeTrim(match[2]) : '';
    var hasTrailingMinus = match[3] === '-';

    if (!pageRange) {
        throw new Error(__("Invalid page range"));
    }

    // Validate and format the page range
    var rangeParts = pageRange.split('-');
    if (rangeParts.length > 2) {
        throw new Error(__("Invalid page range format"));
    }

    var startPage = parseInt(rangeParts[0], 10);
    var endPage = rangeParts.length > 1 ? parseInt(rangeParts[1], 10) : startPage;

    if (isNaN(startPage) || isNaN(endPage)) {
        throw new Error(__("Invalid page numbers"));
    }

    if (startPage > endPage) {
        throw new Error(__("Start page cannot be greater than end page"));
    }

    if (endPage > totalPages) {
        throw new Error(__("Page range exceeds document's total pages"));
    }

    // Format the output
    var formattedPageRange = startPage + (endPage > startPage ? '-' + endPage : '');
    var formattedLabel = label ? '(' + label + ')' : '';
    var formattedInput = formattedPageRange + formattedLabel + (hasTrailingMinus ? '-' : '');

    return formattedInput;
}

function validateInputs() {
    var _pglist = w.tabs.mainPanel.editGroup.pageRangeEntries.text.split(/\n/);
    globalValidatedInputs = [];
    var hasErrors = false;
    var errorMessages = [];
    var seenPageRanges = {};

    for (var i = 0; i < _pglist.length; i++) {
        try {
            var validatedInput = validateAndFormatInput(_pglist[i]);
            var pageRange = extractPageRange(validatedInput);
            
            if (seenPageRanges[pageRange]) {
                errorMessages.push(__("Warning in line") + " " + (i + 1) + ": " + __("Duplicate page range") + " '" + pageRange + "'. " + __("Only the first instance will be exported."));
                continue;
            }
            
            seenPageRanges[pageRange] = true;
            globalValidatedInputs.push(validatedInput);
        } catch (error) {
            errorMessages.push(__("Error in line") + " " + (i + 1) + ": " + error.message);
            hasErrors = true;
        }
    }

    if (errorMessages.length > 0) {
        alert(errorMessages.join("\n"));
    }
    // We only return false if there were actual errors, not warnings
    return !hasErrors;
}

function extractPageRange(input) {
    // Extracts the page range from the formatted input
    var match = input.match(/^(\d+(?:-\d+)?)/);
    return match ? match[1] : '';
}

// Global key handler: Enter in settings panel returns to main panel (with validation)
w.onKeyDown = function(k) {
    try {
        var key = k && k.keyName ? String(k.keyName) : '';
        if (key === 'Enter' || key === 'Return') {
            if (w.tabs && w.tabs.settingsPanel && w.tabs.settingsPanel.visible) {
                // Trigger the settings button logic to validate and switch back
                if (w.mainGroup && w.mainGroup.buttonGroup && w.mainGroup.buttonGroup.settingsButton && typeof w.mainGroup.buttonGroup.settingsButton.onClick === 'function') {
                    w.mainGroup.buttonGroup.settingsButton.onClick();
                }
            }
        }
    } catch (e) {
        Logger.error('Key handler error', e);
    }
}

function exportPDFs(validatedInputs) {
    Logger.info("Starting PDF export for " + validatedInputs.length + " ranges");
    var _path = getFormattedPath(settings.folder, d.fullName.path);
    var _PDFexportPreset = app.pdfExportPresets.item(settings.exportPreset);
    var localizedPageIdentifier = getLocalizedPageIdentifier();
    
    var hasAsyncExport = (typeof d.asynchronousExportFile === "function");
    Logger.debug("Using " + (hasAsyncExport ? "asynchronous" : "synchronous") + " export");

    var isBuiltIn = _PDFexportPreset.name.match(/^\[.*\]$/);
    var tempPreset;
    if (isBuiltIn) {
        Logger.debug("Using built-in preset: " + _PDFexportPreset.name);
        tempPreset = _PDFexportPreset.duplicate();
        tempPreset.name = "Temp_" + _PDFexportPreset.name.replace(/[\[\]]/g, '');
    }
    
    app.scriptPreferences.enableRedraw = false;
    
    // Objekte für Export-Verwaltung
    var exportQueue = [];
    var failedExports = [];
    var totalExports = validatedInputs.length;
    var successfulExports = 0;

    // Erstelle die Export-Queue
    for (var i = 0; i < validatedInputs.length; i++) {
        var line = validatedInputs[i];
        var overwrite = line.charAt(line.length - 1) === "-";
        var match = line.match(/^([^\(]+?)(?:\(([^)]+)\))?(-?)$/);
        
        if (match) {
            var _pageRange = safeTrim(match[1]);
            var _label = match[2] ? safeTrim(match[2]) : "";
            var safeLabel = _label ? sanitizeFileNameSegment(_label) : "";
            var identifier = _label ? _label : _pageRange;

            var version;
            if (overwrite) {
                version = getFileVersion(_path, String(d.name).replace(/\..+$/, ''), identifier) - 1; 
                if (version < 1) version = 1;
            } else {
                version = getFileVersion(_path, String(d.name).replace(/\..+$/, ''), identifier);
            }

            var _versionLabel = '_v' + version;
            var _firstnumber = parseInt(_pageRange.replace(/(^\d+)(.+$)/i, '$1'), 10);
            
            var fileName = String(d.name).replace(/\..+$/, '') + localizedPageIdentifier + 
                         _pageRange + (safeLabel ? "_" + safeLabel : "") + _versionLabel + '.pdf';
            
            var filePath = createSafePath(_path, fileName);
            Logger.debug("Queue export: " + filePath + " (Page range: " + _pageRange + ")");

            // Speichere Export-Informationen
            var exportInfo = {
                fileName: fileName,
                filePath: filePath,
                pageRange: _pageRange,
                firstNumber: _firstnumber,
                label: _label,
                originalInput: line
            };
            exportQueue[exportQueue.length] = exportInfo;
        }
    }

    // Funktion für den Export eines einzelnen PDF
    function exportSinglePDF(exportInfo) {
        try {
            Logger.debug("Exporting: " + exportInfo.filePath);
            
            // Stelle sicher, dass der Zielordner existiert
            var exportFolder = new Folder(File(exportInfo.filePath).parent);
            if (!exportFolder.exists) {
                Logger.debug("Creating export folder: " + exportFolder.fsName);
                var success = exportFolder.create();
                if (!success) {
                    throw new Error("Konnte Zielordner nicht erstellen: " + exportFolder.fsName);
                }
            }
            
            // Prüfe Schreibrechte
            var testFile = new File(exportFolder.fsName + getPathSeparator() + ".write_test");
            var canWrite = testFile.open("w");
            if (canWrite) {
                testFile.close();
                testFile.remove();
            } else {
                throw new Error("Keine Schreibrechte im Zielordner: " + exportFolder.fsName);
            }
            
            // Setze die PDF-Export-Einstellungen
            app.pdfExportPreferences.pageRange = exportInfo.pageRange;
            
            if (isBuiltIn) {
                if (exportInfo.firstNumber % 2 === 0) {
                    tempPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_FACING;
                } else {
                    tempPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_COVER_PAGE;
                }
            } else {
                if (exportInfo.firstNumber % 2 === 0) {
                    _PDFexportPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_FACING;
                } else {
                    _PDFexportPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_COVER_PAGE;
                }
            }
            
            // --- Export ---
            if (hasAsyncExport) {
                // Moderner, nicht blockierender Export
                d.asynchronousExportFile(ExportFormat.PDF_TYPE, File(exportInfo.filePath), false,
                    isBuiltIn ? tempPreset : _PDFexportPreset);
            } else {
                // Ältere InDesign-Version (z. B. CS6): synchroner Export
                d.exportFile(ExportFormat.PDF_TYPE, File(exportInfo.filePath), false,
                    isBuiltIn ? tempPreset : _PDFexportPreset);
            }
            Logger.info("Successfully exported: " + exportInfo.fileName);
            successfulExports++;
            return true;
        } catch (e) {
            Logger.error("Export failed: " + exportInfo.filePath, e);
            return false;
        }
    }

    // Erster Export-Durchgang
    for (var i = 0; i < exportQueue.length; i++) {
        if (!exportSinglePDF(exportQueue[i])) {
            failedExports[failedExports.length] = exportQueue[i];
        }
    }

    // Retry-Logik für fehlgeschlagene Exporte
    if (failedExports.length > 0) {
        Logger.warn(failedExports.length + " files failed to export, showing retry dialog");
        var failedFileNames = [];
        for (var i = 0; i < failedExports.length; i++) {
            failedFileNames[failedFileNames.length] = failedExports[i].fileName;
        }
        var shouldRetry = showRetryDialog(failedFileNames);
        
        while (shouldRetry && failedExports.length > 0) {
            var currentFailures = [];
            // Kopiere failedExports in currentFailures
            for (var i = 0; i < failedExports.length; i++) {
                currentFailures[currentFailures.length] = failedExports[i];
            }
            failedExports = [];  // Reset für den nächsten Durchgang
            
            Logger.info("Retrying export for " + currentFailures.length + " failed files");
            
            // Versuche jedes fehlgeschlagene PDF erneut zu exportieren
            for (var i = 0; i < currentFailures.length; i++) {
                if (!exportSinglePDF(currentFailures[i])) {
                    failedExports[failedExports.length] = currentFailures[i];
                }
            }
            
            // Wenn noch Fehler übrig sind, frage nach weiterem Retry
            if (failedExports.length > 0) {
                Logger.warn(failedExports.length + " files still failed to export");
                failedFileNames = [];
                for (var i = 0; i < failedExports.length; i++) {
                    failedFileNames[failedFileNames.length] = failedExports[i].fileName;
                }
                shouldRetry = showRetryDialog(failedFileNames);
            }
        }
    }

    app.scriptPreferences.enableRedraw = true;
    Logger.info("Export completed: " + successfulExports + "/" + totalExports + " files exported successfully");

    if (tempPreset) {
        tempPreset.remove();
    }
}

function showRetryDialog(failedExports) {
    var retryDialog = new Window("dialog", __("Export failed"));
    retryDialog.orientation = "column";
    retryDialog.alignChildren = ["fill", "top"];
    retryDialog.spacing = 10;
    retryDialog.margins = 16;

    var messageGroup = retryDialog.add("group");
    messageGroup.orientation = "column";
    messageGroup.alignChildren = ["left", "top"];
    messageGroup.spacing = 10;
    messageGroup.margins = [0, 0, 0, 10];
    
    var messageText = messageGroup.add("statictext", undefined, 
        __("The following files could not be exported:"), 
        {multiline: true});
    messageText.preferredSize.width = 400;
    
    var listGroup = retryDialog.add("panel");
    listGroup.orientation = "column";
    listGroup.alignChildren = ["left", "top"];
    listGroup.spacing = 5;
    listGroup.margins = 16;
    
    var scrollGroup = listGroup.add("group");
    scrollGroup.maximumSize.height = 200;
    var fileList = scrollGroup.add("listbox", undefined, failedExports);
    fileList.preferredSize.width = 350;
    fileList.preferredSize.height = 200;

    var helpText = retryDialog.add("statictext", undefined, 
        __("Please close these files in other programs and click 'Retry' to export again."), 
        {multiline: true});
    helpText.preferredSize.width = 400;

    var buttonGroup = retryDialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignChildren = ["center", "center"];
    buttonGroup.spacing = 10;
    
    var retryButton = buttonGroup.add("button", undefined, __("Retry"));
    var cancelButton = buttonGroup.add("button", undefined, __("Cancel"));
    
    retryButton.onClick = function() {
        retryDialog.close(1);
    }
    
    cancelButton.onClick = function() {
        retryDialog.close(0);
    }
    
    return retryDialog.show();
}

// Main execution block

app.doScript(function() {
    var result = w.show();
    if (result == 1) { // The window was closed with the OK button
            try {
                exportPDFs(globalValidatedInputs);
                // Save the settings after the export
                settings.folder = w.tabs.settingsPanel.folderGroup.folderInput.text;
                settings.exportPreset = w.tabs.settingsPanel.presetGroup.presetDropdown.selection.text;
                settings.pageRanges = globalValidatedInputs.join('\n');
                // The window position and size was already updated by onclose-event in settings
            } catch (e) {
                Logger.userError(__("An error occurred during the export:"), e);
            }
    } 
    saveSettings(d, settings);
    unloadDialog(w);
    app.pdfExportPreferences.viewPDF = _backupViewPDF;

}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT, "exportPageRanges");

// Close logger at the end of the script
Logger.close();

/* ################## END MAIN ################## */



/* ################ AUXILIARY ############## */

// Function to format and resolve paths
function getFormattedPath(inputPath, basePath) {
    try {
        Logger.debug("Formatting path - Input: " + inputPath + ", Base: " + basePath);
        // Mac-spezifische Pfadbehandlung
        if ($.os.indexOf("Mac") > -1) {
            var path = inputPath;
            
            // Normalize backslashes to POSIX for macOS
            path = String(path).replace(/\\/g, '/');
            
            // Wenn es kein absoluter Pfad ist, kombiniere mit basePath
            if (!path.match(/^\//)) {
                // Stelle sicher, dass basePath mit Slash endet
                basePath = basePath.replace(/\/?$/, '/');
                // Stelle sicher, dass inputPath nicht mit Slash beginnt
                path = path.replace(/^\//, '');
                path = basePath + path;
            }
            
            // Normalisiere mehrfache Slashes
            path = path.replace(/\/+/g, '/');
            
            Logger.debug("Formatted Mac path: " + path);
            return new Folder(path);
        } else {
            // Windows branch – einfachere und sicherere Implementierung für CS6
            if (/^[a-zA-Z]:/.test(inputPath)) {
                var result = new Folder(inputPath.replace(/[\/\\]+$/, '')); // strip trailing slash
                Logger.debug("Formatted Windows absolute path: " + result.fsName);
                return result;
            } else {
                var result = new Folder(basePath + '/' + inputPath.replace(/^[\/]+/, ''));
                Logger.debug("Formatted Windows relative path: " + result.fsName);
                return result;
            }
        }
    } catch(e) {
        Logger.error("Error in path handling", e);
        
        // Fallback: Verwende einfache Pfadkombination
        Logger.warn("Using fallback path handling");
        return new Folder(basePath + "/" + String(inputPath).replace(/^\//, ''));
    }
}

function formatPath(path) {
    if ($.os.indexOf("Mac") > -1) {
        // macOS path formatting
        return path.replace(/^\//,"");
    } else {
        // Windows path formatting as URI-path i.e. "/C/path"
        return path.replace(/([A-Z]+)(:)/g,"/$1").replace(/\\/g, '/').replace(/\/$/, "");
    }
}


// Helper function to escape special characters in regex
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


function decodeURIComponentSafe(str) {
    return str.replace(/%20/g, ' ').replace(/%([0-9A-F]{2})/g, function(match, p1) {
        return String.fromCharCode('0x' + p1);
    });
}


// Function to save settings to document label
function saveSettings(doc, settings) {
    try {
        Logger.debug("Saving settings to document");
        var windowBoundsString = settings.windowBounds ? 
            [settings.windowBounds.x, settings.windowBounds.y, settings.windowBounds.width, settings.windowBounds.height].join(',') : '';
        var settingsString = SCRIPT_VERSION + "|||" + 
            settings.pageRanges + "|||" + 
            settings.folder + "|||" + 
            settings.exportPreset + "|||" +
            windowBoundsString + "|||" +
            (settings.pageRangePrefix || __("Pages"));
        doc.insertLabel("script_exportPageRangeSettings", settingsString);
        Logger.debug("Settings saved successfully");
    } catch (e) {
        Logger.error("Error saving settings", e);
    }
}

// Function to load settings from document label
function loadSettings(doc) {
    try {
        Logger.debug("Loading settings from document");
        var settingsString = doc.extractLabel("script_exportPageRangeSettings");
        if (settingsString) {
            var settingsArray = settingsString.split("|||");
            
            function isValidSettingsArray(arr) {
                if (arr.length < 5) return false;
                if (!isValidVersion(arr[0])) return false;
                if (typeof arr[1] !== 'string' || !isValidPageRanges(arr[1])) return false;
                if (typeof arr[2] !== 'string' || typeof arr[3] !== 'string') return false;
                if (!isValidWindowBounds(arr[4])) return false;
                return true;
            }

            function isValidVersion(str) {
                var parts = str.split('.');
                if (parts.length !== 2) return false;
                return !isNaN(parseFloat(parts[0])) && !isNaN(parseFloat(parts[1]));
            }

            function isValidPageRanges(str) {
                return str === '' || /^\d/.test(str);
            }

            function isValidWindowBounds(str) {
                if (str === '') return true;
                var parts = str.split(',');
                if (parts.length !== 4) return false;
                for (var i = 0; i < parts.length; i++) {
                    if (isNaN(parseFloat(parts[i]))) return false;
                }
                return true;
            }
            
            function isValidExportPreset(preset) {
                return app.pdfExportPresets.itemByName(preset).isValid;
            }
            
            // check if scriptversion is still same and fits the settings OR check if old settings are still useable
            if (settingsArray[0] === SCRIPT_VERSION || isValidSettingsArray(settingsArray)) { 
                var windowBounds = undefined;
                if (settingsArray[4]) {
                    var boundsArray = settingsArray[4].split(',');
                    if (boundsArray.length === 4) {
                        windowBounds = {
                            x: Number(boundsArray[0]),
                            y: Number(boundsArray[1]),
                            width: Number(boundsArray[2]),
                            height: Number(boundsArray[3])
                        };
                    }
                }
                // test if the folder from settings still exists
                var folder = getFormattedPath(settingsArray[2], doc.fullName.path);
                if (!folder.exists) {
                    settingsArray[2] = "/"; // Reset to default if folder doesn't exist
                    Logger.warn("Workfolder from last settings not available, using default");
                    alert(__("Workfolder from last settings not available. Reverting to default folder."));
                }
                var exportPreset = settingsArray[3];
                if (!isValidExportPreset(exportPreset)) {
                    exportPreset = "[PDF/X-4:2008]"; // Reset to default if preset doesn't exist
                    Logger.warn("Export preset from last settings not available, using default");
                    alert(__("Export preset from last settings not available. Reverting to default preset."));
                }
                
                Logger.info("Settings loaded successfully");
                return {
                    pageRanges: settingsArray[1] || "",
                    folder: settingsArray[2],
                    exportPreset: exportPreset,
                    windowBounds: windowBounds,
                    pageRangePrefix: sanitizeFileNameSegment(settingsArray[5] || __("Pages"))
                };
            } else {
                Logger.warn("Incompatible settings found, using defaults");
                alert(__("Incompatible settings were found. Using default settings."));
            }
        } else {
            Logger.info("No saved settings found, using defaults");
        }
    } catch (e) {
        Logger.error("Error loading settings", e);
    }
    return false;
}

function unloadDialog(dialog) {
    if (dialog && dialog.isValid) {
        try {
            dialog.destroy();
            dialog = null;
            delete dialog;
        } catch (e) {
            Logger.error("Error while unloading dialog", e);
        }
    }
}

// Mac-sichere Pfadzusammensetzung
function createSafePath(basePath, fileName) {
    try {
        if ($.os.indexOf("Mac") > -1) {
            var exportFolder = new Folder(basePath);
            if (!exportFolder.exists) {
                exportFolder.create();
            }
            // Nutze File-Objekt für korrekte Pfadkonvertierung
            var tempFile = new File(exportFolder + "/" + fileName);
            var result = decodeURI(tempFile.fsName);
            Logger.debug("Created safe Mac path: " + result);
            return result;
        }
        var result = new File(new Folder(basePath).fsName + getPathSeparator() + fileName).fsName;
        Logger.debug("Created safe Windows path: " + result);
        return result;
    } catch (e) {
        Logger.error("Error creating safe path", e);
        return basePath + getPathSeparator() + fileName;
    }
}

// Funktion zur Validierung des Pfads im Eingabefeld
function validateFolderPath() {
    try {
        if (d.isValid && d.saved && d.fullName !== null) {
            var inputPath = w.tabs.settingsPanel.folderGroup.folderInput.text;
            
            // Einfachere Folder-Erstellung für CS6
            var formattedFolder;
            try {
                formattedFolder = getFormattedPath(inputPath, d.fullName.path);
            } catch(e) {
                Logger.userError(__("Path validation error:"), e);
                return false;
            }

            // Prüfe ob Zielordner existiert
            if (!formattedFolder.exists) {
                try {
                    // Versuche den Ordner zu erstellen
                    Logger.debug("Attempting to create folder: " + formattedFolder.fsName);
                    var success = formattedFolder.create();
                    if (!success) {
                        Logger.warn("Failed to create directory: " + formattedFolder.fsName);
                        alert(__("Cannot create directories at the specified path:") + " " + formattedFolder.fsName);
                        return false;
                    }
                } catch(e) {
                    Logger.error("Error creating directory", e);
                    alert(__("Cannot create directories at the specified path:") + " " + formattedFolder.fsName);
                    return false;
                }
            }
            
            // Prüfe Schreibrechte über eine einfachere Methode
            try {
                var testFile = new File(formattedFolder.fsName + getPathSeparator() + ".write_test");
                var canWrite = testFile.open("w");
                if (canWrite) {
                    testFile.close();
                    testFile.remove();
                } else {
                    Logger.warn("No write permissions in folder: " + formattedFolder.fsName);
                    alert(__("No write permissions in the specified folder:") + " " + formattedFolder.fsName);
                    return false;
                }
            } catch(e) {
                Logger.error("Error testing write permissions", e);
                alert(__("No write permissions in the specified folder:") + " " + formattedFolder.fsName);
                return false;
            }
            
            return true;
        } else {
            Logger.warn("File has no directory path until saved");
            alert(__("File has no directory path until saved."));
            return false;
        }
    } catch (e) {
        Logger.userError(__("Path validation error:"), e);
        return false;
    }
}


