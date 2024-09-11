var WINDOWWIDTH = 400;
var WINDOWHEIGHT = 400;
var test = 101;
resizeTo(WINDOWWIDTH, WINDOWHEIGHT);
moveTo(screen.width / 2 - WINDOWWIDTH / 2, screen.height / 2 - WINDOWHEIGHT / 2)

var shell = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

var pathBtn = document.getElementById('choosePath')
pathBtn.onclick = openD;

function openD() {    
    
    var fileName = createFolerDialogFile();  
    shell.Run("%TMP%/" + fileName,1,true);         
}

function createFolerDialogFile() {
    try {
       var fileName = "popUptmp.hta";       

        var body = "<html>";

        body += "<head>"
        body += "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css' integrity='sha384-xOolHFLEh07PJGoPkLv1IbcEPTNtaed2xpHsD9ESMhqIYd0nLMwNLD69Npy4HI+N' crossorigin='anonymous'></link>"
        body += "</head>"

        body += "<body id='popUpBody'>";
        body += "<div id='file-system'></div>"
        body += "</body>"
        
        body += "<script>"
        body += "document.title='Выбрать директорию';resizeTo(800,500);moveTo(screen.width/2-400,screen.height/2-250);";
        body += "var body = document.getElementById('popUpBody');"
        body += "body.innerHTML = '<h1>Выбор папки</h1>';"

          
        
        body += "</script>"

        body += "</html>"

        var tempFile = fso.GetSpecialFolder(2) + "\\" + fileName;
        var file = fso.CreateTextFile(tempFile, true);
        file.Write(body);
        file.Close();

        return fileName;
       //return shell.Exec("mshta.exe \"file://" + tempFile + "\"")
    } catch (e) {        
        alert("Ошибка: " + e.message);
        return null;
    }
}

/*test part. put into createFolerDialogFile aftrer test*/
var drives = [];
var ds = new Enumerator(fso.Drives);
    for (; !ds.atEnd(); ds.moveNext()) {
        var drive = ds.item();
        drives.push({
            name: drive.DriveLetter + ":\\",
            type: drive.DriveType,
            shareName: drive.ShareName
           
        });
       
        
             
    }
    
createFileSystemHTML("C:\\");    


function createFileSystemHTML(path) {
    var data = getFileSystemData(path)
    
    var fileSystemContainer = document.getElementById('file-system');   
    fileSystemContainer.innerHTML = "<div class='folder'><input  type='button' class='btn btn-primary plusBtn' value='../' onclick='goUpFolder(\"" + data.path + "\", \"" + data.type +"\")'/> " + data.name + "</div>";
    
    if (data.children.length) {
        for (var i=0; i < data.children.length; i++ ) {
            fileSystemContainer.innerHTML += "<div class='subFolder'><input  type='button' class='btn btn-secondary plusBtn' value='+' onclick='createFileSystemHTML(\"" + data.children[i].path.join("\\\\") + "\")'/>" + data.children[i].name + "</div>";            
        }
        
    }  
}


function getFileSystemData(path) {    
    try {       
        var folder = fso.GetFolder(path); 

        return parseFolder(folder);
    } catch (e) {
        alert("Ошибка: " + e.message);
        return null;
    }
}

function parseFolder(folder) {    
    
    var data = {        
        name: folder.Name || folder.Drive, 
        path: folder.Path,
        type: folder.Name ? 1 : 0,
        children: []
        
    };    
    
    
    // Добавляем подпапки
    var subFolders = new Enumerator(folder.subFolders);
    for (; !subFolders.atEnd(); subFolders.moveNext()) {
        var subFolder = subFolders.item();
        data.children.push({
            name: subFolder.Name,            
            path: subFolder.Path.split('\\')                                    
        });
        
    }
       

    return data;
}

function goUpFolder(path, type) {
    if (type == 1) {
        createFileSystemHTML(path + '\\../')
    }
}
/* end test part. put into createFolerDialogFile aftrer test*/

