var PATHTOSAMPLE = "\\\\nas\\ОГТ\\МТЦ";
//var PATHTOSAMPLE = "E:\\Downloads";
//var PATHTOORDERS = ["E:\\Downloads"];
var PATHTOORDERS = [];
var date = new Date;
var oderFolder = fso.GetFolder("\\\\nas\\Запросы\\");
oderFolder += date.getFullYear();
alert (oderFolder)
var subOderFolders = new Enumerator(oderFolder.subFolders);
for (; !subOderFolders.atEnd(); subOderFolders.moveNext()) {
    
}


var WINDOWWIDTH = 400;
var WINDOWHEIGHT = 400;
var test = 101;
resizeTo(WINDOWWIDTH, WINDOWHEIGHT);
moveTo(screen.width / 2 - WINDOWWIDTH / 2, screen.height / 2 - WINDOWHEIGHT / 2)

try {
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
} catch(e) {
    alert("Ошибка: " + e.message);
    close();        
}

var pathOfInstall = shell.Environment("User")('hta001path');
var pathBtn = document.getElementById('choosePath');
var targetPath = document.getElementById('targetPath');
var crtBtn = document.getElementById('create');

if (pathOfInstall) {
    targetPath.value = pathOfInstall;
} else {
    targetPath.value = "";
}

pathBtn.onclick = openD;
crtBtn.onclick = createFolder;


function openD() {
    var fileName = createFolerDialogFile();  
    shell.Run("%TMP%/" + fileName,1,true);
    pathOfInstall = shell.Environment("User")('hta001path');
    targetPath.value = pathOfInstall;         
}

function createFolerDialogFile() {
    try {
       var fileName = "popUptmp.hta";       

        var body = "<html>";

        body += "<head>";
        body += "<hta:application border='thin' caption='yes' maximizeButton='no' minimizeButton='yes' showInTaskbar='yes' singleInstance='yes' sysMenu='yes' windowState='normal' scroll='no'/>";
        body += "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css' integrity='sha384-xOolHFLEh07PJGoPkLv1IbcEPTNtaed2xpHsD9ESMhqIYd0nLMwNLD69Npy4HI+N' crossorigin='anonymous'></link>";
        body += "<style>";
            body += "#file-system {height: 23em; overflow-y: scroll;}";
            body += ".subFolder {margin-left: 1em; margin-top: 0.1em;}";
            body += ".folder {font-weight: bold; width: 100%;}";
            body += ".plusBtn {width: 2em;}";
            body += ".reqBtn {width: 8em; margin: 1em; float: right;}";
        body += "</style>";
        body += "</head>";

        body += "<body>";
        body += "<div class='container'><h4>Выберите папку установки</h4></div>";
        body += "<div class='container' id='file-system'></div>";
        body += "<div class='container'><input type='button' class='btn btn-secondary reqBtn' value='cancel' id='cnlBtn'/><input type='button' class='btn btn-primary reqBtn' value='ok' id='okBtn'/></div>";
        body += "</body>";
        
        body += "<script>";
        body += "document.title='Выберите папку установки';resizeTo(600,500);moveTo(screen.width/2-300,screen.height/2-250);";
        body += "var shell = new ActiveXObject('WScript.Shell');";
        body += "var fso = new ActiveXObject('Scripting.FileSystemObject');";
        body += "var pathOfInstall = shell.Environment('User')('hta001path');";
        body += "var okBtn = document.getElementById('okBtn');";
        body += "var clBtn = document.getElementById('cnlBtn');";

        body += "clBtn.onclick = function cl() {close();};";
        body += "okBtn.onclick = function ok() {shell.Environment('User')('hta001path') = pathOfInstall; close();};";

        body += " if (pathOfInstall) {createFileSystemHTML(pathOfInstall.split('\\\\').join('\\\\\\\\'));} else {createFileSystemHTML('C:\\\\');}";       
            
        body += "function createFileSystemHTML(path, data) {"; 
        body += "pathOfInstall = path;";       
        body += "var fileSystemContainer = document.getElementById('file-system');";
        body += "if (!data) {";
        body += "var data = getFileSystemData(path);";
        body += "fileSystemContainer.innerHTML = \"<div class='folder'><input  type='button' class='btn btn-primary plusBtn' value='../' onclick='goUpFolder(\\\"\" + data.path.join(\"\\\\\\\\\") + \"\\\",\" + data.type + \")'/>\" + data.path.join(\"\\\\\") + \"</div>\"; ";
        body += "} else {  fileSystemContainer.innerHTML = \"<div class='folder'>Computer</div>\";}";
        body += "if (data.children.length) {";
        body += "for (var i=0; i < data.children.length; i++ ) {";
        body += "fileSystemContainer.innerHTML += \"<div class='subFolder'><input  type='button' class='btn btn-secondary plusBtn' value='+' onclick='createFileSystemHTML(\\\"\" + data.children[i].path.join(\"\\\\\\\\\") + \"\\\")'/>\" + data.children[i].name + \"</div>\";";
        body += "}}}";
        
        body += "function getFileSystemData(path) { try { var folder = fso.GetFolder(path); return parseFolder(folder); } catch (e) { alert('Ошибка: ' + e.message); return null;}}";
        
        body += "function parseFolder(folder) {";
        body += "var data = { name: folder.Name || folder.Drive, path: folder.Path.split('\\\\'), type: folder.Name ? 1 : 0, children: [] };";
        body += "var subFolders = new Enumerator(folder.subFolders); for (; !subFolders.atEnd(); subFolders.moveNext()) { var subFolder = subFolders.item(); data.children.push({ name: subFolder.Name, path: subFolder.Path.split('\\\\') });}";
        body += "return data;}";
        
        body += "function goUpFolder(path, type) {";
        body += "if (type == 1) { createFileSystemHTML(path + '\\\\../'); }";
        body += " if (type == 0) { var data = { children: [] };";
        body += "var ds = new Enumerator(fso.Drives);";
        body += "for (; !ds.atEnd(); ds.moveNext()) { var drive = ds.item(); data.children.push({ path: (drive.DriveLetter + ':\\\\').split('\\\\'), name: drive.ShareName || drive.DriveLetter });}";
        body += "createFileSystemHTML('',data); }}";
        body += "</script>";
        
        body += "</html>";

        var tempFile = fso.GetSpecialFolder(2) + "\\" + fileName;
        var file = fso.CreateTextFile(tempFile, true);
        file.Write(body);
        file.Close();

        return fileName;       
    } catch (e) {        
        alert("Ошибка: " + e.message);
        return null;
    }
}

function createFolder() {
    var styleSheet = document.styleSheets[0];    
    styleSheet.addRule('*', 'cursor: wait;');
    
    
    
    try {   
    
    if (!targetPath.value) {
        openD();
    } 
    var orderNum = document.getElementById('orderNum');
    var customer = document.getElementById('customer').value;
    var name = document.getElementById('name').value;
    var bpNum = document.getElementById('bpNum').value;

    if ((/[A-ZА-Я0-9\.\_\-()]+/gi).test(orderNum.value)) {        
        var newPath = targetPath.value;
        var newName = orderNum.value;       
        if (customer) { newName += ' ' + customer; } 
        if (name) { newName += ' ' + name; } 
        if (bpNum) { newName += ' ' + bpNum; } 
        newPath += '\\'+ newName;
       
        //create new folders
        fso.createFolder(newPath);        
        fso.createFolder(fso.buildPath(newPath, '01 ПРЕДРАСЧЕТ'));
        fso.createFolder(fso.buildPath(fso.buildPath(newPath, '01 ПРЕДРАСЧЕТ'), 'ОМТС'));
        fso.createFolder(fso.buildPath(fso.buildPath(newPath, '01 ПРЕДРАСЧЕТ'), 'НОРМИРОВАНИЕ'));
        fso.createFolder(fso.buildPath(newPath, '02 ЧЕРТЕЖИ'));
        fso.createFolder(fso.buildPath(fso.buildPath(newPath, '02 ЧЕРТЕЖИ'), 'ЭСКИЗЫ'));
        fso.createFolder(fso.buildPath(newPath, '03 ТЗ'));
        fso.createFolder(fso.buildPath(newPath, '04 ТЕХПРОЦЕСС'));
        fso.createFolder(fso.buildPath(newPath, '05 ЗАПУСК'));
        fso.createFolder(fso.buildPath(fso.buildPath(newPath, '05 ЗАПУСК'), 'ОМТС'));
        fso.createFolder(fso.buildPath(fso.buildPath(newPath, '05 ЗАПУСК'), 'НОРМИРОВАНИЕ'));
        
        //copy files
        var folder = fso.GetFolder(PATHTOSAMPLE);
        var files = new Enumerator(folder.files);        
        for (; !files.atEnd(); files.moveNext()) {
            var file = files.item();
            if (file.Name.indexOf('Образец полный') !== -1) {
                var command = 'powershell.exe -Command "Copy-Item -Path \'' + file.path + '\' -Destination \'' + fso.buildPath(newPath, '01 ПРЕДРАСЧЕТ') + '\'"';                 
                shell.Run(command, 0, true);                                
                fso.GetFile(fso.buildPath(newPath, '01 ПРЕДРАСЧЕТ') + "\\" + file.Name).name = newName + '.xlsm';                
            }
            if (file.Name.indexOf('ТЗ Шаблон') !== -1) {
                var command = 'powershell.exe -Command "Copy-Item -Path \'' + file.path + '\' -Destination \'' + fso.buildPath(newPath, '03 ТЗ') + '\'"';                 
                shell.Run(command, 0, true);                                
            }
            if (file.Name.indexOf('Маршруты Шаблон') !== -1) {
                var command = 'powershell.exe -Command "Copy-Item -Path \'' + file.path + '\' -Destination \'' + fso.buildPath(newPath, '04 ТЕХПРОЦЕСС') + '\'"';                 
                shell.Run(command, 0, true);                                
            }
        }
        
        //create links
        var shortcut = shell.CreateShortcut(fso.buildPath(newPath, '06 Запрос.lnk'));            
        for (var i = 0; i < PATHTOORDERS.length; i++) {
            folder = fso.GetFolder(PATHTOORDERS[i]);
            var subFolders = new Enumerator(folder.subFolders);
            for (; !subFolders.atEnd(); subFolders.moveNext()) {
                var j = subFolders.item();                
                if (j.path.indexOf(orderNum.value) !== -1) {
                    shortcut.targetPath = j.path;
                    shortcut.save();                    
                    break;
                }
            }

        }
        if (!shortcut.targetPath) {alert('запрос с таким номером не найден, ярлык не добвален')}
        alert('папка создана');

    } else {
        alert('введите номер запроса\nдопустимые символы: A-Z А-Я 0-9 \. \_ \- ( )');
    }

    } catch(e) {
        alert("Ошибка: " + e.message);        
    }
    styleSheet.removeRule(styleSheet.rules.length - 1);            
            
}


