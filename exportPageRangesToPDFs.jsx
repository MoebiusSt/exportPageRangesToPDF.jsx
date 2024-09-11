/*******************************************************************************
 * Script Name: exportPageRangesToPDF.jsx
 * Description: This InDesign script exports page ranges to separate PDFs (and manages files a bit).
 *              
 * 
 * Features:
 * - export multiple page ranges to individual PDFs.
 * - specify export-folder and PDF-preset to use.
 * - manually input page ranges a labels/identifiers for each range.
 * - retrieve a list of page ranges from former exported PDFs from that directory
 * - retrieve a list of page ranges from the active documents sections and section markers.
 * – retrieve the section(s) you are currently viewing for a quick versioned export of what you are working on.
 * - handles odd/even page starts in PDF-view-settings (cover sheet YES/NO)
 * - automatic versioning for exported PDFs, unless overridden.
 * - saves settings to a document script label, will start with last settings and position.
 * 
 * User input schema:
 * Each line you enter in the text area should follow this format. The label in brackets is optional, don't use commas:
 *     Page or PageRange(label)
 *
 * Example:
 *     1
 *     5-7(Introduction)
 *     8-15
 *     16
 *     17-20(Chapter 2)
 * 
 * The exported PDF files will be named according to this schema:
 *     [DocumentName]_pages_[PageRange]_[Label-from-brackets-if-given]_v[Version].pdf
 *
 * Example:
 *     /MyDocument_pages_5-7_Introduction_v1.pdf 
 *
 * The version-number will auto increment if it finds an existing pdf file with the same name, i.e. from _v1 to _v2.
 * In case you want to overwrite the last pdf file version you can suppress the version incrementation by adding a minus to the end of a line. 
 *
 * Example:
 *     1-
 *     5-7(Introduction)-
 *     17-20
 *
 *     If "MyDocument_pages_1_v1.pdf" and "MyDocument_pages_5-7_Introduction_v1.pdf" already exists it will overwrite those PDFs during export of those page ranges and will not create version _v2 pdfs. The other page range of 17-20 will be exported with version-incrementation to the next higher version. Of course if there was no prior pdf of any of those files, it will create v1-files.
 * 
 * Author: [Stephan Möbius]
 * Original idea: [Jongware]
 * Multi language L10N-function: Marc Autret, equalizer.js, 2012.
 *
 * Changelog and Script on GitHub:
 * https://github.com/MoebiusSt/exportPageRangesToPDF
 *
 ******************************************************************************/
 
 #targetengine 'exportPageRangesToPDF'
 
//======================================
// <L10N> :: FRENCH_LOCALE :: GERMAN_LOCALE :: SPANISH_LOCALE :: RUSSIAN_LOCALE :: ARABIC_LOCALE
//======================================
// Export Page Ranges :: Exporter des plages de pages :: Seitenbereiche exportieren :: Exportar rangos de páginas :: Экспорт диапазонов страниц :: تصدير نطاقات الصفحات
// Pages :: Pages :: Seiten :: Páginas :: Страницы :: صفحات
// page range :: plage de pages :: Seitenbereich :: rango de páginas :: диапазон страниц :: نطاق الصفحات
// PDF Folder :: Dossier PDF :: PDF-Ordner :: Carpeta PDF :: Папка PDF :: مجلد PDF
// Get from Directory :: Obtenir du répertoire :: Aus Verzeichnis abrufen :: Obtener del directorio :: Получить из каталога :: الحصول من الدليل
// Get from Sections :: Obtenir des sections :: Aus Abschnitten abrufen :: Obtener de secciones :: Получить из разделов :: الحصول من الأقسام
// Sections from this Spread :: Sections de ce planche :: Abschnitte dieses Druckbogens :: Secciones de este pliego :: Разделы этого разворота :: أقسام هذا الانتشار
// Select PDF folder :: Sélectionner le dossier PDF :: PDF-Ordner auswählen :: Seleccionar carpeta PDF :: Выберите папку PDF :: حدد مجلد PDF
// PDF Preset :: Préréglage PDF :: PDF-Vorgabe :: Ajuste preestablecido de PDF :: Предустановка PDF :: الإعداد المسبق لـ PDF
// Error in button click handler :: Erreur dans le gestionnaire de clic de bouton :: Fehler im Button-Klick-Handler :: Error en el controlador de clic del botón :: Ошибка в обработчике нажатия кнопки :: خطأ في معالج النقر على الزر
// Error in getFromDirectory :: Erreur dans getFromDirectory :: Fehler in getFromDirectory :: Error en getFromDirectory :: Ошибка в getFromDirectory :: خطأ في getFromDirectory
// No matching page sections detected from files. Files must match the script's naming scheme :: Aucune section de page correspondante détectée dans les fichiers. Les fichiers doivent correspondre au schéma de nommage du script :: Keine Seitenbereiche aus den Dateien erkannt. Dateien müssen dem Namensschema des Skriptes entsprechen :: No se detectaron secciones de página coincidentes en los archivos. Los archivos deben coincidir con el esquema de nomenclatura del script :: Не обнаружены соответствующие разделы страниц в файлах. Файлы должны соответствовать схеме именования скрипта :: لم يتم اكتشاف أقسام صفحة مطابقة من الملفات. يجب أن تتطابق الملفات مع مخطط التسمية الخاص بالبرنامج النصي
// Incompatible settings were found. Using default settings. :: Des paramètres incompatibles ont été trouvés. Utilisation des paramètres par défaut. :: Inkompatible Einstellungen wurden gefunden. Es werden die Standardeinstellungen verwendet. :: Se encontraron configuraciones incompatibles. Usando configuración predeterminada. :: Обнаружены несовместимые настройки. Используются настройки по умолчанию. :: تم العثور على إعدادات غير متوافقة. استخدام الإعدادات الافتراضية.
// Workfolder from last settings not available. Reverting to default folder. :: Le dossier de travail des derniers paramètres n'est pas disponible. Retour au dossier par défaut. :: Arbeitsordner aus letzten Einstellungen nicht verfügbar. Zurück zum Standardordner. :: La carpeta de trabajo de la última configuración no está disponible. Volviendo a la carpeta predeterminada. :: Рабочая папка из последних настроек недоступна. Возврат к папке по умолчанию. :: مجلد العمل من الإعدادات الأخيرة غير متوفر. العودة إلى المجلد الافتراضي.
// Export preset from last settings not available. Reverting to default preset. :: Le préréglage d'exportation des derniers paramètres n'est pas disponible. Retour au préréglage par défaut. :: Exportvorgabe aus letzten Einstellungen nicht verfügbar. Zurück zur Standardvorgabe. :: El ajuste preestablecido de exportación de la última configuración no está disponible. Volviendo al ajuste preestablecido predeterminado. :: Экспортный пресет из последних настроек недоступен. Возврат к пресету по умолчанию. :: الإعداد المسبق للتصدير من الإعدادات الأخيرة غير متوفر. العودة إلى الإعداد المسبق الافتراضي.
// This script requires the InDesign setting 'General > Numbering > Absolute Page Numbering'. It will apply this setting. Do you want to proceed? :: Ce script nécessite le paramètre InDesign 'Général > Numérotation > Numérotation absolue des pages'. Il appliquera ce paramètre. Voulez-vous continuer ? :: Dieses Skript erfordert die InDesign-Einstellung 'Allgemein > Nummerierung > Absolute Seitennummerierung'. Es wird diese Einstellung anwenden. Möchten Sie fortfahren? :: Este script requiere la configuración de InDesign 'General > Numeración > Numeración de página absoluta'. Aplicará esta configuración. ¿Desea continuar? :: Этот скрипт требует настройки InDesign 'Общие > Нумерация > Абсолютная нумерация страниц'. Он применит эту настройку. Хотите продолжить? :: يتطلب هذا البرنامج النصي إعداد InDesign 'عام > الترقيم > ترقيم الصفحات المطلق'. سيتم تطبيق هذا الإعداد. هل تريد المتابعة؟
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
// </L10N>

var SCRIPT_VERSION = "2.01"; // Date: 2024-09-11

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

// Function for trimming that can handle both strings and null/undefined
function safeTrim(str) {
    return str == null ? '' : String(str).trim();
}


function checkAndSetAbsolutePageNumbering() {
    if (app.generalPreferences.pageNumbering !== PageNumberingOptions.ABSOLUTE) {
        // Prompt the user with the option to continue or cancel
        var response = confirm(__("This script requires the InDesign setting 'General > Numbering > Absolute Page Numbering'. It will apply this setting. Do you want to proceed?"), true, __("Setting required"));
        
        if (!response) {
            exit();
        }
        app.generalPreferences.pageNumbering = PageNumberingOptions.ABSOLUTE;
    }
}


function getFileVersion(path, baseName, identifier) {
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
    return highestVersion + 1;
}

function getLocalizedPageIdentifier() {
    return "_" + __("Pages") + "_";
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
        var folder = new Folder(_path);
        var localizedPageIdentifier = getLocalizedPageIdentifier();
        var docName = d.name.replace(/\..+$/, '');
        var filePattern = escapeRegExp(docName + localizedPageIdentifier) + ".+_v\\d+\\.pdf$";
        
        var files = folder.getFiles(function(file) {
            return decodeURI(file.name).match(new RegExp(filePattern, "i"));
        });
        
        if (files.length === 0) {
            return __("No matching page sections detected from files. Files must match the script's naming scheme") + ": \r\n\r\n" +
                   __("Example") + ":\r\n" +
                   "\"" + d.name.replace(/\..+$/, '') + localizedPageIdentifier + "11-12_Advertorial_v02.pdf" + "\""; 
        }

        var pageRanges = {};
        for (var i = 0; i < files.length; i++) {
            var file = files[i];
			var fileName = file.name;
            var regex = new RegExp("^.+" + escapeRegExp(localizedPageIdentifier) + "(\\d+-?\\d*)_?(.*)_v(\\d+)\\.pdf$", "i");
            var match = fileName.match(regex);
			
            if (match) {
                var pageRange = safeTrim(match[1]);
                var label = fixEncoding(decodeURIComponent(safeTrim(match[2])));
                
                // If label is empty, keep it empty; use pageRange as a key but don't display it as a label
                var identifier = label || pageRange;
                
                var version = parseInt(match[3], 10);
                
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
        
        return safeTrim(result);
    } catch (error) {
        alert(__("Error in getFromDirectory") + ": " + error.message);
        return __("An unexpected error occurred in the script.");
    }
}

function getSectionPages() {
    var sections = d.sections;
    var sectionInfo = [];
    
    for (var i = 0; i < sections.length; i++) {
        var section = sections[i];
        var start = section.pageStart.documentOffset + 1;
        var end = start + section.length - 1;
        
        var sectionName = createSectionName(section);
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
    
    return result.join("\n");
}

function createSectionName(section) {
    var parts = [];
    if (section.includeSectionPrefix && section.sectionPrefix !== "") {
        parts.push(section.sectionPrefix);
    }
    if (section.marker !== "") {
        parts.push(section.marker);
    }
    return parts.join("-");
}

function formatSectionRange(start, end, sectionName) {
    var range = start === end ? start.toString() : start + "-" + end;
    return sectionName ? range + "(" + sectionName + ")" : range;
}


function getActiveSection() {
    var activeWindow = app.activeWindow;

    if (!activeWindow) {
        alert(__("No active window detected. Please ensure a document is open and try again."));
        return "";
    }
    var activeSpread = activeWindow.activeSpread;
    var activePage = activeWindow.activePage;
    // Check whether we are on a MasterPage
    if ((activeSpread && activeSpread.parent instanceof MasterSpread) ||
        (activePage && activePage.parent instanceof MasterSpread)) {
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
                spreadSections.push(formatSectionRange(
                    section.pageStart.documentOffset + 1, 
                    section.pageStart.documentOffset + section.length, 
                    createSectionName(section)
                ));
                seenSections[section.id] = true;
            }
        }
    } catch (e) {
        alert(__("An error occurred while retrieving section information: ") + e.message);
        return "";
    }
    return spreadSections.join("\n");
}



/* ################## MAIN ################## */
var d = app.activeDocument;

var settings = loadSettings(d) || {
    /* DEFAULT SETTINGS ... Note: if we change the schema of the settings then remember to adjust the loadSettings validation! */
    pageRanges: "",
    folder: "/",
    exportPreset: "[PDF/X-4:2008]",
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
    text: '" + __("Export Page Ranges") + "', \
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
        settingsButton: Button { \
            text: '" + __("Settings") + "', \
            properties: {name: 'Settings'}, \
            alignment: ['right', 'top'], \
            margins:0 \
        }, \
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
            presetGroup: Group { \
                orientation: 'column', \
                preferredSize: [160,20],\
                alignChildren: ['left', 'center'], \
                presetLabel: StaticText { text: '" + __("PDF Preset") + ":' }, margins:[0,15,0,0], \
                presetDropdown: DropDownList { preferredSize: [160,20], alignment: ['fill', 'center'] } \
            } \
        }, \
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


// Resize event handler
w.onResizing = w.onResize = function() {
    this.layout.resize();
}

// onClose event to save the window position and size
w.onClose = function() {
    settings.windowBounds = this.bounds;
};

// Define behavior for the SETTINGS button  
w.mainGroup.settingsButton.onClick = function ()	{
    with (w.tabs) {
        if (settingsPanel.visible == false) {	
            settingsPanel.visible = true;	
            mainPanel.visible = false; 
            this.active = true;
        }
        else { 
            settingsPanel.visible = false;	
            mainPanel.visible = true;
            this.active = false;                
        }
    }
}; 

function withButtonWorkaround(button, fn) {// higher order function for a workaround to unhighlight lit buttons
    return function() {
        button.active = true;
        try {
            fn.apply(this, arguments);
        } catch (error) {
            alert(__("Error in button click handler") + ": " + error.message);
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
                var relativePath = selectedFolder.fsName.replace(decodeURI(d.filePath.fsName), '');
                w.tabs.settingsPanel.folderGroup.folderInput.text = settings.folder = relativePath + "\\";
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
    
    // Check whether the input starts with a number
    if (!/^\d/.test(trimmedInput)) {
        throw new Error(__("Input must start with a number"));
    }

    // Split the input into page area and label parts
    var match = trimmedInput.match(/^(\d+(?:\s*-\s*\d+)?)\s*(?:\(([^)]*)\))?\s*(-)?$/);
    if (!match) {
        throw new Error(__("Invalid input format"));
    }
    
    var pageRange = match[1] ? match[1].replace(/\s+/g, '') : '';  // Entfernt alle Leerzeichen im Seitenbereich
    var label = match[2] ? safeTrim(match[2]) : '';
    var hasTrailingMinus = match[3] === '-';

    if (!pageRange) {
        throw new Error(__("Invalid page range"));
    }

    // Validate and format the page area
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
                // If we have already seen this page area, we warn the user
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


function exportPDFs(validatedInputs) {
    var _path = getFormattedPath(settings.folder, d.fullName.path);
    var _PDFexportPreset = app.pdfExportPresets.item(settings.exportPreset);
    var localizedPageIdentifier = getLocalizedPageIdentifier();
    
    var isBuiltIn = _PDFexportPreset.name.match(/^\[.*\]$/);
    var tempPreset;
    if (isBuiltIn) {
        tempPreset = _PDFexportPreset.duplicate();
        tempPreset.name = "Temp_" + _PDFexportPreset.name.replace(/[\[\]]/g, '');
    }
    
    app.scriptPreferences.enableRedraw = false;
    
    var failedExports = [];
    var totalExports = validatedInputs.length;
    var successfulExports = 0;

    for (var i = 0; i < validatedInputs.length; i++) {
        var line = validatedInputs[i];
        var overwrite = line.charAt(line.length - 1) === "-";
        var match = line.match(/^([^\(]+?)(?:\(([^)]+)\))?(-?)$/);
        
        if (match) {
            var _pageRange = safeTrim(match[1]);
            var _label = match[2] ? safeTrim(match[2]) : "";
            
            var identifier = _label ? _label : _pageRange;

            var version;
            if (overwrite) {
                version = getFileVersion(_path, String(d.name).replace(/\..+$/, ''), identifier) - 1; 
                if (version < 1) version = 1;
            } else {
                version = getFileVersion(_path, String(d.name).replace(/\..+$/, ''), identifier);
            }

            var _versionLabel = '_v' + version;

            app.pdfExportPreferences.pageRange = _pageRange;
            var _firstnumber = Number(_pageRange.replace(/(^\d+)(.+$)/i, '$1'));
            
            if (isBuiltIn) {
                if (_firstnumber.isEven()) {
                    tempPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_FACING;
                } else {
                    tempPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_COVER_PAGE;
                }
            } else {
                if (_firstnumber.isEven()) {
                    _PDFexportPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_FACING;
                } else {
                    _PDFexportPreset.pdfPageLayout = PageLayoutOptions.TWO_UP_COVER_PAGE;
                }
            }
            
            var fileName = String(d.name).replace(/\..+$/, '') + localizedPageIdentifier + _pageRange + (_label ? "_" + _label : "") + _versionLabel + '.pdf';
            var fileNameAndPath = _path + "/" + fileName;
            
            try {
                d.asynchronousExportFile(ExportFormat.PDF_TYPE, File(fileNameAndPath), false, isBuiltIn ? tempPreset : _PDFexportPreset);
                successfulExports++;
            } catch (e) {
                failedExports.push(fileName);
            }
        }
    }

    app.scriptPreferences.enableRedraw = true;

    if (tempPreset) {
        tempPreset.remove();
    }

    if (failedExports.length > 0) {
        var failedCount = failedExports.length;
        alert(__("%1 of %2 files could not be exported. In most cases, the cause is files that are already open. Please close:").replace("%1", failedCount).replace("%2", totalExports) + "\n\n" + failedExports.join("\n"));
    }
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
                alert(__("An error occurred during the export:") + " " + e.message);
            }
    } 
    saveSettings(d, settings);
    unloadDialog(w);
    app.pdfExportPreferences.viewPDF = _backupViewPDF;

}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT, "exportPageRanges");

/* ################## END MAIN ################## */



/* ################ AUXILIARY ############## */

// Function to format and resolve paths
function getFormattedPath(inputPath, basePath) {
    if ($.os.indexOf("Mac") > -1) {
        // macOS path handling
        if (inputPath.charAt(0) === '/') {
            return Folder(inputPath);
        } else {
            return Folder(basePath + '/' + inputPath);
        }
    } else {
        // Windows path handling
        if (inputPath.match(/^[a-zA-Z]:/)) {
            return Folder(formatPath(inputPath) + '/' + inputPath);
        } else {
            return Folder(basePath + inputPath);
        }
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


function decodeURIComponent(str) {
    return str.replace(/%20/g, ' ').replace(/%([0-9A-F]{2})/g, function(match, p1) {
        return String.fromCharCode('0x' + p1);
    });
}

function fixEncoding(str) {
    return str ? str.replace(/Ã¤/g, "ä")
                    .replace(/Ã¶/g, "ö")
                    .replace(/Ã¼/g, "ü")
                    .replace(/Ã/g, "ß")
                    .replace(/Ã/g, "Ä")
                    .replace(/Ã/g, "Ö")
                    .replace(/Ã/g, "Ü")
                    .replace(/Ã©/g, "é")
                    .replace(/Ã /g, "à")
                    .replace(/Ã¨/g, "è")
                    .replace(/Ã¹/g, "ù")
                    .replace(/Ãª/g, "ê")
                    .replace(/Ã¢/g, "â")
                    .replace(/Â/g, "")
                    : str;
}


// Function to save settings to document label
function saveSettings(doc, settings) {
    var windowBoundsString = settings.windowBounds ? 
        [settings.windowBounds.x, settings.windowBounds.y, settings.windowBounds.width, settings.windowBounds.height].join(',') : '';
    var settingsString = SCRIPT_VERSION + "|||" + 
        settings.pageRanges + "|||" + 
        settings.folder + "|||" + 
        settings.exportPreset + "|||" +
        windowBoundsString;
    doc.insertLabel("script_exportPageRangeSettings", settingsString);
}

// Function to load settings from document label
function loadSettings(doc) {
    var settingsString = doc.extractLabel("script_exportPageRangeSettings");
    if (settingsString) {
        var settingsArray = settingsString.split("|||");
        
        function isValidSettingsArray(arr) {
            if (arr.length !== 5) return false;
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
                alert(__("Workfolder from last settings not available. Reverting to default folder."));
            }
            var exportPreset = settingsArray[3];
            if (!isValidExportPreset(exportPreset)) {
                exportPreset = "[PDF/X-4:2008]"; // Reset to default if preset doesn't exist
                alert(__("Export preset from last settings not available. Reverting to default preset."));
            }
            return {
                pageRanges: settingsArray[1] || "",
                folder: settingsArray[2],
                exportPreset: exportPreset,
                windowBounds: windowBounds
            };
        } else {
            alert(__("Incompatible settings were found. Using default settings."));
        }
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
            alert("Error while unloading dialog: " + e);
        }
    }
}