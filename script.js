var WINDOWWIDTH = 400;
var WINDOWHEIGHT = 400;

resizeTo(WINDOWWIDTH, WINDOWHEIGHT);
moveTo(screen.width / 2 - WINDOWWIDTH / 2, screen.height / 2 - WINDOWHEIGHT / 2)

var pathBtn = document.getElementById('choosePath')
pathBtn.onclick = openFolerDialog;

function openFolerDialog() {
    var shell = new ActiveXObject("WScript.Shell");
    var script = "<script>document.title='������� ����������'; var WINDOWWIDTH = 800; var WINDOWHEIGHT = 500; resizeTo(WINDOWWIDTH, WINDOWHEIGHT); moveTo(screen.width / 2 - WINDOWWIDTH / 2, screen.height / 2 - WINDOWHEIGHT / 2)</script>";
    var body = "<hta:application id='selectFolder' border='thin' caption='yes' maximizeButton='no' minimizeButton='no' showInTaskbar='yes' singleInstance='yes' sysMenu='yes' windowState='normal' scroll='no'/><h2>������� �����</h2>"    
    
    shell.Exec("mshta.exe \"about:" + body + script +"\"")
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

   