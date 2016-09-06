function SelectAll() {
    for (var i=0;i<document.delall.chkid.length;i++) {
        var e=document.delall.chkid[i];
        e.checked=!e.checked;
    }
}

$(document).ready(function(){
    $('#add').click(function(){
    if($("#addtable").hasClass("hidden")){
        $("#addtable").removeClass("hidden");
        $("#table").addClass("hidden");
        $("#edittable").addClass("hidden");
        $("#pagination").addClass("hidden");
    }else{
        $("#addtable").addClass("hidden");
        $("#table").removeClass("hidden");
        $("#edittable").removeClass("hidden");
        $("#pagination").removeClass("hidden");
    }
    })
});