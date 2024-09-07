var WINDOWWIDTH = 400;
var WINDOWHEIGHT = 400;
var test = 101;
resizeTo(WINDOWWIDTH, WINDOWHEIGHT);
moveTo(screen.width / 2 - WINDOWWIDTH / 2, screen.height / 2 - WINDOWHEIGHT / 2)

var shell = new ActiveXObject("WScript.Shell");

var pathBtn = document.getElementById('choosePath')
pathBtn.onclick = openD;

function openD() {    
    
    var fileName = createFolerDialogFile();  
    shell.Run("%TMP%/" + fileName,1,true);         
}

function createFolerDialogFile() {
    try {
       var fileName = "popUptmp.hta"
        var fso = new ActiveXObject("Scripting.FileSystemObject");

        var body = "<html>";

        body += "<head>"
        body += "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css' integrity='sha384-xOolHFLEh07PJGoPkLv1IbcEPTNtaed2xpHsD9ESMhqIYd0nLMwNLD69Npy4HI+N' crossorigin='anonymous'></link>"
        body += "</head>"

        body += "<body id='popUpBody'>";
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
