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
    //script += "document.title='������� ����������';resizeTo(800,500);moveTo(screen.width/2-400,screen.height/2-250);";    
   // script += "</script>"

    //var body = "<body id='popUpBody'></body>"; 
    //alert("mshta.exe \"about:" + body + script +"\"")       
    //shell.Exec("mshta.exe \"about:" + body + script +"\"")
    //shell.Exec("mshta.exe \"about:file://" + "\"")
}   




function selectFolder() {
        try {
            var objExcel = new ActiveXObject("Excel.Application");
            var objDialog = objExcel.FileDialog(4); // 4 ������������� msoFileDialogFolderPicker

            objDialog.Title = "�������� �����";
            objDialog.AllowMultiSelect = false;
            objDialog.InitialFileName = "C:\\Users"; // ��������� ��������� ����� ��� ������������ �����

            // ������������� ���� HTA �� �������� ����
            setForegroundWindow();

            if (objDialog.Show() == -1) { // ���� ������������ ������ �����
                var strFolderPath = objDialog.SelectedItems(1);
                alert("�� ������� �����: " + strFolderPath);
            } else {
                alert("����� ����� �������.");
            }

            objExcel.Quit();
        } catch (e) {
            alert("������: " + e.message);
        }
}

   