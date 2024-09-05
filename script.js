var WINDOWWIDTH = 400;
var WINDOWHEIGHT = 400;
var test = 101;
resizeTo(WINDOWWIDTH, WINDOWHEIGHT);
moveTo(screen.width / 2 - WINDOWWIDTH / 2, screen.height / 2 - WINDOWHEIGHT / 2)

var pathBtn = document.getElementById('choosePath')
pathBtn.onclick = openFolerDialog;

function openFolerDialog() {
    var shell = new ActiveXObject("WScript.Shell");    
   // var script = "<script></script>"
    //script += "document.title='Выбрать директорию';resizeTo(800,500);moveTo(screen.width/2-400,screen.height/2-250);";    
   // script += "</script>"

    //var body = "<body id='popUpBody'></body>"; 
    //alert("mshta.exe \"about:" + body + script +"\"")       
    //shell.Exec("mshta.exe \"about:" + body + script +"\"")
    //shell.Exec("mshta.exe \"about:file://" + "\"")
}   




function selectFolder() {
        try {
            var objExcel = new ActiveXObject("Excel.Application");
            var objDialog = objExcel.FileDialog(4); // 4 соответствует msoFileDialogFolderPicker

            objDialog.Title = "Выберите папку";
            objDialog.AllowMultiSelect = false;
            objDialog.InitialFileName = "C:\\Users"; // Указываем начальную папку без завершающего слэша

            // Устанавливаем окно HTA на передний план
            setForegroundWindow();

            if (objDialog.Show() == -1) { // Если пользователь выбрал папку
                var strFolderPath = objDialog.SelectedItems(1);
                alert("Вы выбрали папку: " + strFolderPath);
            } else {
                alert("Выбор папки отменен.");
            }

            objExcel.Quit();
        } catch (e) {
            alert("Ошибка: " + e.message);
        }
}

   