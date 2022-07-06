

let dataModel = {
    siteUrl: _spPageContextInfo.webAbsoluteUrl,
    isEdit: false,
    isAddFn: false,
    currentEditId: null,
    loginUserEmail: null,
};

$(function () {
    eventFnInit();
    getUserInfomation((userInfo)=>{
        console.log(`当前登录人信息：${userInfo}`);
        $("#loginUser").html(userInfo.userName);
        getDataList((d) => {
            renderHtml(handlerDatas(d));
        });
    });
});

/**
 * 所有点击事件初始化 
 */
function eventFnInit() {
    // 新增按钮点击
    $("#addBtn").click(() => {
        $("#urlType").val('');
        $("#urlDisplayName").val('');
        $("#urlAddress").val('');
        $("#editModalLabel").html('新增');
        dataModel.isAddFn = true;
        $('#editModal').modal('show');
    });
    // 修改按钮点击
    $("body").on('click', '.editBtn', (e) => {
        let id = $(e.target).attr('data-id');
        dataModel.currentEditId = id;
        $("#editModalLabel").html('修改');
        $("#urlType").val($(e.target).parent().find('a').eq(0).attr('type-title'));
        $("#urlDisplayName").val($(e.target).parent().find('a').eq(0).attr('title'));
        $("#urlAddress").val($(e.target).parent().find('a').eq(0).attr('href'));
        dataModel.isAddFn = false;
        $('#editModal').modal('show');
    });
    // 编辑按钮点击
    $("#editBtn").click(() => {
        dataModel.isEdit = !dataModel.isEdit;
        if (dataModel.isEdit) {
            $(".delBtn,.editBtn").show();
            $("#editBtn").html(`<i class="fa fa-edit"></i> 取消编辑`);
        } else {
            $(".delBtn,.editBtn").hide();
            $("#editBtn").html(`<i class="fa fa-edit"></i> 编辑`);
        }
    });
    // 新增/修改事件
    $("#submitBtn").click(() => {
        if (dataModel.isAddFn) {
            let clientContext = new SP.ClientContext(dataModel.siteUrl);
            let workflowList = clientContext.get_web().get_lists().getByTitle("Quick Links");
            let updateListItem = workflowList.addItem(new SP.ListItemCreationInformation());
            updateListItem.set_item("Type_x0020_Name", $("#urlType").val());
            updateListItem.set_item("Display_x0020_Title", $("#urlDisplayName").val());
            updateListItem.set_item("Url", $("#urlAddress").val());
            let createBy = SP.FieldUserValue.fromUser(dataModel.loginUserEmail);
            updateListItem.set_item('Owner', createBy);
            updateListItem.update();
            clientContext.load(updateListItem);
            clientContext.executeQueryAsync(function () {
                $('#editModal').modal('hide');
                location.reload();
            }, function () {
                alert("Service error, submit button click!");
            });
        } else {
            // JSRequest.EnsureSetup();
            let clientContext = new SP.ClientContext(dataModel.siteUrl);
            let updateList = clientContext.get_web().get_lists().getByTitle("Quick Links");
            let updateListItem = updateList.getItemById(dataModel.currentEditId);
            updateListItem.set_item("Type_x0020_Name", $("#urlType").val());
            updateListItem.set_item("Display_x0020_Title", $("#urlDisplayName").val());
            updateListItem.set_item("Url", $("#urlAddress").val());
            updateListItem.update();
            clientContext.load(updateListItem);
            clientContext.executeQueryAsync(function () {
                $('#editModal').modal('hide');
                location.reload();
            }, function () { alert("Service Technical Error!"); });
        }

    });
    // 删除按钮点击
    $("body").on('click', '.delBtn', (e) => {
        let id = $(e.target).attr('data-id');
        let clientContext = new SP.ClientContext(dataModel.siteUrl);
        let orderList = clientContext.get_web().get_lists().getByTitle("Quick Links");
        let camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml('<View><Query><Where><Eq><FieldRef Name="ID" /><Value Type="Text">' + id + '</Value></Eq></Where></Query></View>');
        let collListItem = orderList.getItems(camlQuery);
        clientContext.load(collListItem);
        clientContext.executeQueryAsync(function () {
            let itemCount = collListItem.get_count();
            for (let i = itemCount - 1; i >= 0; i--) {
                let oListItem = collListItem.itemAt(i);
                oListItem.deleteObject();
            }
            clientContext.executeQueryAsync(function () {
                getDataList((d) => {
                    renderHtml(handlerDatas(d));
                });
            }, function () { alert("Can not delete order, draft button click!"); });
        }, function () { alert("Can not access order list, draft button click!"); });
    });
}

/**
 * 获取当前登录人信息
 */
function getUserInfomation(callback) {
    let current_usergroups = [];
    let current_user = null;
    let usergroups;
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
        let clientContext = SP.ClientContext.get_current();
        let oWeb = clientContext.get_web();
        current_user = oWeb.get_currentUser();
        usergroups = current_user.get_groups();
        clientContext.load(current_user);
        clientContext.load(usergroups);
        clientContext.executeQueryAsync(function () {
            let userName = current_user.get_title();
            let userEmail = current_user.get_email();
            let groupEnumerator = usergroups.getEnumerator();
            while (groupEnumerator.moveNext()) {
                let oGroup = groupEnumerator.get_current();
                current_usergroups.push(oGroup.get_title());
            }
            dataModel.loginUserEmail = userEmail;
            let userInfo = {
                userName,
                userEmail,
                current_usergroups
            }
            if(callback){
                callback(userInfo);
            }
        }, function () {
            alert("Can not get current user information!");
        });
    });
}

// 查询数据
function getDataList(callback) {
    $("#preloader").fadeIn("slow");
    let apiUrl = dataModel.siteUrl + "/_api/web/lists/GetByTitle('Quick Links')/items?$select=" +
        "ID," +
        "Owner/Name," +
        "Owner/EMail," +
        "Type_x0020_Name," +
        "Display_x0020_Title," +
        "Url"+
        "&$expand=Owner&$filter=Owner/EMail eq '" + dataModel.loginUserEmail + "'";
    $.ajax({
        url: apiUrl,
        method: "GET",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function (data) {
            if (callback) {
                callback(data.d.results);
            }
        },
        error: function (e) {
            console.log(e);
            alert("Can not access Quick Links List!");
        }
    });
}
// 渲染页面
function renderHtml(processGuidData) {
    let html = "";
    for (item in processGuidData) {
        let typeTitle = processGuidData[item]['Type_x0020_Name'];
        let childrenList = processGuidData[item]['children'];
        let childrenHtml = '';
        childrenList.forEach((childItem, idx) => {
            childrenHtml += ` <li><a type-title='${typeTitle}' href="${childItem['Url']}" title='${childItem['Display_x0020_Title']}' target="_blank">${childItem['Display_x0020_Title']}</a><span data-id='${childItem['ID']}' class="delBtn">X</span><span class="editBtn"  data-id='${childItem['ID']}'>修改</span></li>`
        });
        html += `<div class="type-item">
        <div class="item-title">
            <h2><i class='fa fa-bookmark'></i> ${typeTitle}</h2>
        </div>
        <div class="item-detail">
            <ul>
                ${childrenHtml}    
            </ul>
        </div>
    </div>`
    }
    $("#mainList").html(html);
    if (dataModel.isEdit) {
        $(".delBtn,.editBtn").show();
    } else {
        $(".delBtn,.editBtn").hide();
    }
}

// 根据Type_x0020_Name重新构造对象
function handlerDatas(arr) {
    let obj = {};
    arr.forEach((item, index) => {
        let { Type_x0020_Name } = item;
        if (!obj[Type_x0020_Name]) {
            obj[Type_x0020_Name] = {
                Type_x0020_Name,
                children: []
            }
        }
        obj[Type_x0020_Name].children.push(item);
    });
    // return Object.values(obj);
    return obj;
}