
/*test part. put into createFolerDialogFile aftrer test*/

var fso = new ActiveXObject("Scripting.FileSystemObject");
    
createFileSystemHTML("C:\\");    


function createFileSystemHTML(path, data) {
    var fileSystemContainer = document.getElementById('file-system'); 
    if (!data) {
        var data = getFileSystemData(path);
        
        fileSystemContainer.innerHTML = "<div class='folder'><input  type='button' class='btn btn-primary plusBtn' value='../' onclick='goUpFolder(\"" + data.path.join("\\\\") + "\"," + data.type + ")'/>" + data.name + "</div>"; 
    } else {
        fileSystemContainer.innerHTML = "<div class='folder'>Computer</div>"; 
    }
    
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
        path: folder.Path.split('\\'),
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
        createFileSystemHTML(path + '\\../');
    }
    if (type == 0) {
        var data = {
            children: []
        };
        
        var ds = new Enumerator(fso.Drives);
        for (; !ds.atEnd(); ds.moveNext()) {
            var drive = ds.item();
            data.children.push({
                path: (drive.DriveLetter + ":\\").split('\\'),
                name: drive.ShareName || drive.DriveLetter
           
            });  
        }
        createFileSystemHTML("",data);
    }
    
}
/* end test part. put into createFolerDialogFile aftrer test*/