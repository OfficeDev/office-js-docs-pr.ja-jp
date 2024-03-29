---
title: テキスト エディターを使用して Microsoft Project 用の作業ウィンドウ アドインを初めて作成する
description: Office アドイン用 Yeoman ジェネレーターを使用して、Project Standard 2013、Project Professional 2013 以降のバージョンの作業ウィンドウ アドインを作成します。
ms.date: 07/10/2020
ms.localizationpriority: medium
ms.openlocfilehash: 69353b94da05dd0a8cfd6347beb7c5b3f4fd81e0
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66959029"
---
# <a name="create-your-first-task-pane-add-in-for-microsoft-project-by-using-a-text-editor"></a>テキスト エディターを使用して Microsoft Project 用の作業ウィンドウ アドインを初めて作成する

Office アドイン用 Yeoman ジェネレーターを使用して、Project Standard 2013、Project Professional 2013 以降のバージョンの作業ウィンドウ アドインを作成できます。この記事では、ファイル共有上の HTML ファイルを指す XML マニフェストを使用する単純なアドインを作成する方法について説明します。 Project OM Test サンプル アドインは、アドインにオブジェクト モデルを使用するいくつかの JavaScript 関数をテストします。プロジェクトの **セキュリティ センター** を使用してマニフェスト ファイルを含むファイル共有を登録したら、リボンの [ **プロジェクト** ] タブから作業ウィンドウ アドインを開くことができます。 Project 2013の セキュリティ センターを使用して、マニフェスト ファイルを含むファイル共有を登録した後は、作業ウィンドウ アドインをリボンの [ プロジェクト] タブから開くことができます (この記事のサンプル コードは、Microsoft Corporation の Arvind Iyer によるテスト アプリケーションに基づくものです)。

Project では、他の Office クライアントが使用するのと同じアドイン マニフェスト スキーマと、同じ JavaScript API の大部分が使用されます。 この記事に記載されているアドインの完全なコードは、Project 2013 SDK ダウンロードのサブディレクトリ `Samples\Apps` で提供されています。

Project OM Test サンプル アドインは、タスクの GUID と、アプリケーションおよびアクティブなプロジェクトのプロパティを取得できます。 Project Professional 2013 で SharePoint ライブラリ内にあるプロジェクトを開くと、このアドインでは、そのプロジェクトの URL を表示できます。

[Project 2013 SDK のダウンロード](https://www.microsoft.com/download/details.aspx?id=30435)には完全なソース コードが含まれています。Project2013SDK.msi に含まれる SDK を展開してインストールしたら、`\Samples\Apps\Copy_to_AppManifests_FileShare` サブディレクトリにマニフェスト ファイルがあり、`\Samples\Apps\Copy_to_AppSource_FileShare` サブディレクトリにソース コードがあることを確認します。

サンプルの JSOMCall.html では、インクルードされる office.js ファイルと project-15.js ファイル内の JavaScript 関数を使用しています。 対応するデバッグ ファイル (office.debug.js および project-15.debug.js) を使用すると、これらの関数を検証できます。

Office アドインでの JavaScript の使用の概要については、「 [Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) について」を参照してください。

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a>手順 1. アドイン マニフェスト ファイルを作成するには

ローカル ディレクトリに XML ファイルを作成します。 この XML ファイルには `OfficeApp` 要素と子要素を記述します。詳細については「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。 たとえば、次の XML を含むJSOM_SimpleOMCalls.xmlという名前のファイルを作成します (要素の GUID 値を `Id` 変更します)。

```XML
<?xml version="1.0" encoding="utf-8"?>
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
              xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
              xsi:type="TaskPaneApp">
     <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
     <Id>93A26520-9414-492F-994B-4983A1C7A607</Id>
     <Version>15.0</Version>
     <ProviderName>Microsoft</ProviderName>
     <DefaultLocale>en-us</DefaultLocale>
     <DisplayName DefaultValue="Project OM Test">
       <Override Locale="fr-fr" Value="Le Project OM Test"/>
     </DisplayName>
     <Description DefaultValue="Test the task pane add-in object model for Project - English (US)">
       <Override Locale="fr-fr" Value="Test the task pane add-in object model for Project - French (France)"/>
     </Description>
     <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
     <Hosts>
       <Host Name="Project"/>
       <Host Name="Workbook"/>
       <Host Name="Document"/>
     </Hosts>
    <DefaultSettings>
       <SourceLocation DefaultValue="\\ServerName\AppSource\JSOMCall.html">
         <Override Locale="fr-fr" Value="\\ServerName\AppSource\JSOMCall.html"/>
       </SourceLocation>
     </DefaultSettings>
     <Permissions>ReadWriteDocument</Permissions>
     <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
       <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
     </IconUrl>
     <AllowSnapshot>true</AllowSnapshot>
   </OfficeApp>
```

Project の場合、要素には `OfficeApp` 属性値を `xsi:type="TaskPaneApp"` 含める必要があります。 要素は `Id` GUID です。 この値は `SourceLocation` 、アドイン HTML ソース ファイルまたは作業ウィンドウで実行される Web アプリケーションのファイル共有パスまたは SharePoint URL である必要があります。 For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).

手順 2. では、JSOM_SimpleOMCalls.xml マニフェストが Project テスト アドインのために指定する HTML ファイルの作成方法を示します。この HTML 内で指定されているボタンは、関連する JavaScript 関数を呼び出します。JavaScript 関数は、この HTML ファイル内に追加したり、別の .js ファイル内に配置したりできます。

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a>手順 2. Project OM Test アドインのソース ファイルを作成するには

1. JSOM_SimpleOMCalls.xml マニフェストの要素で指定された名前の `SourceLocation` HTML ファイルを作成します。

   たとえば、`C:\Project\AppSource`ディレクトリで theJSOMCall.html ファイルを作成します。 単純なテキスト エディターを使用してソース ファイルを作成できますが、特定のドキュメントの種類 (HTML や JavaScript など) で動作し、他の編集補助機能を備えた Visual Studio Code などのツールを使用する方が簡単です。 「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」で説明されている Bing Search の例をまだ行っていない場合は、マニフェストが指定する `\\ServerName\AppSource` ファイル共有を作成する方法が手順 3 で示されています。　

   JSOMCall.html ファイルでは、AJAX 機能に共通のMicrosoftAjax.js ファイルと、Office 2013 アプリケーションのアドイン機能用のOffice.js ファイルが使用されます。

    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <title>Project OM Sample Code</title>
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <script type="text/javascript" src="MicrosoftAjax.js"></script>

            <!-- Use the CDN reference to office.js when deploying your add-in. -->
            <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
            <script type="text/javascript" src="office.js"></script>
            <script type="text/javascript" src="JSOM_Sample.js"></script>
        </head>
        <body>
            <div id="Common_JSOM_API">
                OBJECT MODEL TESTS
            </div>

            <textarea id="text" rows="6" cols="25">This is the text result.</textarea>
        </body>
    </html>
    ```

   この要素は `textarea` 、JavaScript 関数の結果を示すテキスト ボックスを指定します。

   > [!NOTE]
   > Project OM Test サンプルを実行するには、Project 2013 SDK ダウンロードに含まれるファイル (Office.js、Project-15.js、および MicrosoftAjax.js) を JSOMCall.html ファイルと同じディレクトリにコピーします。

   手順 2. では、Project OM Test サンプル アドインが使用する特定の関数のために JSOM Sample.js というファイルを追加します。この後の手順では、JavaScript 関数を呼び出すボタン用にその他の HTML 要素を追加します。

1. JSOM_Sample.js という名前の JavaScript ファイルを、JSOMCall.html ファイルと同じディレクトリ内に作成します。

   次のコードは、Office.js ファイル内のメソッドを使用してアプリケーション コンテキストとドキュメント情報を取得します。 オブジェクトは `text` 、HTML ファイル内の `textarea` コントロールの ID です。

   **\_projDoc** 変数はオブジェクトで`ProjectDocument`初期化されます。 このコードには、いくつかの簡単なエラー処理関数と、アプリケーション コンテキストと `getContextValues` プロジェクト ドキュメント コンテキストのプロパティを取得する関数が含まれています。 Project の JavaScript オブジェクト モデルの詳細については、「[JavaScript API for Office](../reference/javascript-api-for-office.md)」を参照してください。

    ```js
    /*
    * JavaScript functions for the Project OM Test example app
    * in the Project 2013 SDK.
    */

    var _projDoc;
    var _app;
    var taskGuid;
    var resourceGuid;

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready method.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            _projDoc = Office.context.document;
            _app = Office.context;
        });
    }

    function logError(errorText) {
        text.value = "Error in " + errorText;
    }

    function logEventError(erroneousEvent) {
        logError("event " + erroneousEvent);
    }

    function logMethodError(methodName, errorName, errorMessage) {
        logError(methodName + " method.\nError name: " + errorName + "\nMessage: " + errorMessage);
    }

    // . . . Add other JavaScript functions here.

    function getContextValues() {
        getDocumentUrl();
        getDocumentMode();
        getApplicationContentLanguage();
        getApplicationDisplayLanguage();
    }

    function getDocumentUrl() {
        text.value ="Document URL:\n" +_projDoc.url;
    }

    function getDocumentMode() {
        var docMode = _projDoc.mode;
        text.value = text.value + "\n\nDocument mode: " + docMode;
    }

    function getApplicationContentLanguage() {
        text.value = text.value + "\nApp language: " + _app.contentLanguage;
    }

    function getApplicationDisplayLanguage() {
        text.value = text.value + "\nDisplay language: " + _app.displayLanguage;
    }
    ```

   Office.debug.js ファイルの関数の詳細については、「 [Office JavaScript API](../reference/javascript-api-for-office.md)」を参照してください。 たとえば、関数は `getDocumentUrl` 、開いているプロジェクトの URL またはファイル パスを取得します。

1. Office.jsとProject-15.jsで非同期メソッドを呼び出して選択したデータを取得する JavaScript 関数を追加します。

   - たとえば、 `getSelectedDataAsync` 選択したデータの書式設定されていないテキストを取得するOffice.jsの一般的なメソッドです。 詳細については、「 [Document オブジェクト](/javascript/api/office/office.document#office-office-document-getselectedtaskasync-member(1))」を参照してください。

   - Project-15.jsの関数は `getSelectedTaskAsync` 、選択したタスクの GUID を取得します。 同様に `getSelectedResourceAsync` 、この関数は選択したリソースの GUID を取得します。 タスクまたはリソースが選択されていない状態でこれらの関数を呼び出すと、未定義のエラーが発生します。

   - この関数は `getTaskAsync` 、タスク名と割り当てられたリソースの名前を取得します。 タスクが同期された SharePoint タスク リスト内にある場合は、 `getTaskAsync` SharePoint の一覧でタスク ID を取得します。それ以外の場合、SharePoint タスク ID は 0 です。

     > [!NOTE]
     > サンプル コードには、デモ用にバグが含まれています。 未定義の場合 `taskGuid` 、関数は `getTaskAsync` エラーアウトします。有効なタスク GUID を取得し、別のタスクを選択した場合、 `getTaskAsync` 関数は関数によって操作された最新のタスクのデータを `getSelectedTaskAsync` 取得します。
  
   - `getTaskFields`、`getResourceFields`および`getProjectFields`タスクまたは`getProjectFieldAsync`リソースの指定されたフィールドを取得するために複数回`getResourceFieldAsync`呼び出`getTaskFieldAsync`すローカル関数です。 project-15.debug.js ファイルでは、 `ProjectTaskFields` 列挙体と `ProjectResourceFields` 列挙体に、サポートされているフィールドが表示されます。

   - この関数は `getSelectedViewAsync` 、ビューの種類 (project-15.debug.jsの `ProjectViewTypes` 列挙で定義) とビューの名前を取得します。

   - プロジェクトが SharePoint タスクリストと同期されている場合、 `getWSSUrlAsync` 関数は URL とタスクリストの名前を取得します。 プロジェクトが SharePoint タスクリストと同期されていない場合、関数は `getWSSUrlAsync` エラーになります。

     > [!NOTE]
     > SharePoint の URL とタスクリストの名前を取得するには、[ProjectProjectFields](/javascript/api/office/office.projectprojectfields) 列挙体の定数と`WSSList`共`WSSUrl`にメソッドを使用`getProjectFieldAsync`することをお勧めします。

   次のコードの各関数には、`function (asyncResult)` によって指定されている匿名関数が含まれます。これは、非同期の結果を取得するコールバックです。匿名関数の代わりに、複雑なアドインの保守に役立つ名前付き関数を使用できます。

    ```js
    // Get the data in the selected cells of the grid in the active view.
    function getSelectedDataAsync() {
        _projDoc.getSelectedDataAsync(
            Office.CoercionType.Text,
            { ValueFormat: "Formatted" },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded)
                    text.value = asyncResult.value;
                else
                    logMethodError("getSelectedDataAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        );
    }

    // Get the GUID of the selected task.
    function getSelectedTaskAsync() {
        _projDoc.getSelectedTaskAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                taskGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get the GUID of the selected resource.
    function getSelectedResourceAsync() {
        _projDoc.getSelectedResourceAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                resourceGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get data for the specified task.
    function getTaskAsync() {
        if (taskGuid != undefined) {
            _projDoc.getTaskAsync(
                taskGuid,
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        logMethodError("getTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                    } else {
                        var taskInfo = asyncResult.value;
                        var taskOutput = "Task name: " + taskInfo.taskName +
                                         "\nGUID: " + taskGuid +
                                         "\nWSS Id: " + taskInfo.wssTaskId +
                                         "\nResourceNames: " + taskInfo.resourceNames;
                        text.value = taskOutput;
                    }
                }
            );
        } else {
            text.value = 'Task GUID not valid:\n' + taskGuid;
        } 
    }

    // Get additional data for task fields.
    function getTaskFields() {
        text.value = "";

        _projDoc. getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Name: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.ID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "ID: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Start: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Duration,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Duration: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Priority,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Priority: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Notes,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Notes: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        ); 
    }

    // Get data for the specified resource fields.
    function getResourceFields() {
        text.value = "";

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Resource name: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Cost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.StandardRate,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Standard Rate: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualCost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualWork,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Work: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Units,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Units: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );
    }

    // Get the URL and list name of the synchronized SharePoint task list.
    // Recommended: use getProjectField instead.
    function getWSSUrlAsync() {
        _projDoc.getWSSUrlAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = "SharePoint URL:\n" + asyncResult.value.serverUrl
                    + "\nList name: " + asyncResult.value.listName;
            }
            else {
                logMethodError("getWSSUrlAsync", asyncResult.error.name, asyncResult.error.message);
            }
        });
    }

    // Get the type and name of the selected view.
    function getSelectedViewAsync() {
        _projDoc.getSelectedViewAsync(function (asyncResult) {
            text.value = "View type: " + asyncResult.value.viewType
                + "\nName: " + asyncResult.value.viewName;
        });
    }

    // Get information about the active project.
    function getProjectFields() {
        text.value = "";

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Project GUID: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nStart: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Finish,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nFinish: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProject " + errorText);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencyDigits,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nCurrency digits: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbol,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Currency symbol: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbolPosition,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSymbol position: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nProject web app URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSList,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint list: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    ```

1. JavaScript イベント ハンドラー コールバックおよび関数を追加して、タスク選択、リソース選択、およびビュー選択の変更に関するイベント ハンドラーの登録と登録解除を行います。 この関数は `manageEventHandlerAsync` 、 *操作* パラメーターに応じて、指定されたイベント ハンドラーを追加または削除します。 この操作には、次の操作を指定 `addHandlerAsync` できます `removeHandlerAsync`。

   関数`manageResourceEventHandler`は`manageTaskEventHandler``manageViewEventHandler`、*docMethod* パラメーターで指定されているように、イベント ハンドラーを追加または削除できます。

    ```js
    // Task selection changed event handler.
    function onTaskSelectionChanged(eventArgs) {
        text.value = "In task selection change event handler";
    }

    // Resource selection changed event handler.
    function onResourceSelectionChanged(eventArgs) {
        text.value = "In Resource selection changed event handler";
    }

    // View selection changed event handler.
    function onViewSelectionChanged(eventArgs) {
        text.value = "In View selection changed event handler";
    }

    // Add or remove the specified event handler.
    function manageEventHandlerAsync(eventType, handler, operation, onComplete) {
        _projDoc[operation]   //The operation is addHandlerAsync or removeHandlerAsync.
        (
            eventType,
            handler,
            function (asyncResult) {
                if (onComplete) {
                    onComplete(asyncResult, operation);
                } else {
                    var message = "Operation: " + operation;
                    message = message + "\nStatus: " + asyncResult.status + "\n";
                    text.value = message;
                }
            }
        );
    }

    // Write the asyncResult status from the manageEventHandlerAsync function (optional). 
    function onComplete(asyncResult, operation) {
        var message = "In onComplete function for " + operation;
        message = message + "\nStatus: " + asyncResult.status;
        text.value = message;
    }

    // Add or remove a task selection changed event handler.
    function manageTaskEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.TaskSelectionChanged,      // The task selection changed event.
            onTaskSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a resource selection changed event handler.
    function manageResourceEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ResourceSelectionChanged,  // The resource selection changed event.
            onResourceSelectionChanged,                 // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a view selection changed event handler.
    function manageViewEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ViewSelectionChanged,      // The view selection changed event.
            onViewSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }
    ```

1. この HTML ドキュメントの本文に、テストのために JavaScript 関数を呼び出すボタンを追加します。 たとえば、一般的な JSOM API の要素で `div` 、一般的 `getSelectedDataAsync` な関数を呼び出す入力ボタンを追加します。

    ```HTML
    <body>
        <div id="Common_JSOM_API">
        OBJECT MODEL TESTS
        <br /><br />
        <strong>General function:</strong>
        <br />
        <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
            value="getSelectedDataAsync" />
        </div>
        <!--  more code . . .  -->
    ```

1. プロジェクト固有の `div` タスク関数とイベントのボタンを含むセクションを `TaskSelectionChanged` 追加します。

    ```HTML
    <div id="ProjectSpecificTask">
      <br />
      <strong>Project-specific task methods:</strong><br />
      <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
      <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
      <strong>Task selection changed:</strong>
      <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    ```

1. リソースメソッドとイベントのボタンを含むセクションを追加 `div` し、メソッドとイベント、プロジェクトプロパティ、コンテキストプロパティを表示する

    ```HTML
    <div id="ResourceMethods">
      <br />
      <strong>Resource methods:</strong>
      <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
      <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
      <strong>Resource selection changed:</strong>
      <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ViewMethods">
      <br />
      <strong>View method:</strong>
      <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
      <strong>View selection changed:</strong>
      <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ProjectMethods">
      <br />
      <strong>Project properties:</strong>
      <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
    </div>
    <div id="ContextVariables">
      <br />
      <strong>Context properties:</strong>
      <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
    </div>
    ```

1. ボタン要素を書式設定するには、CSS `style` 要素を追加します。 たとえば、要素の子として次を `head` 追加します。

    ```HTML
    <style type="text/css">
        .button-wide
        {
            width: 210px;
            margin-top: 2px;
        }
        .button-narrow
        {
            width: 80px;
            margin-top: 2px;
        }
    </style>
    ```

手順 3. では、Project OM Test アドインの機能をインストールして使用する方法を示します。

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a>手順 3. Project OM Test アドインをインストールして使用するには

1. JSOM SimpleOMCalls.xml マニフェストが含まれているディレクトリに対するファイル共有を作成します。 ファイル共有は、ローカル コンピューター上、またはネットワーク上のアクセス可能なリモート コンピューター上に作成できます。 たとえば、マニフェストがローカル コンピューターのディレクトリにある  `C:\Project\AppManifests` 場合は、次のコマンドを実行します。

    `Net share AppManifests=C:\Project\AppManifests`

1. Project OM Test アドインの HTML および JavaScript ファイルが含まれるディレクトリに対するファイル共有を作成します。 このファイル共有パスは、JSOM SimpleOMCalls.xml マニフェストで指定されているパスに一致するようにしてください。 たとえば、ファイルがローカル コンピューターのディレクトリにある  `C:\Project\AppSource` 場合は、次のコマンドを実行します。

    `net share AppSource=C:\Project\AppSource`

1. Project で、[**Project のオプション**] ダイアログ ボックスを開き、[**セキュリティ センター**]、[**セキュリティ センターの設定**] の順に選択します。

   アドインの登録手順および追加情報については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。

1. **[セキュリティ センター]** ダイアログ ボックスの左側のウィンドウで、**[信頼されているアドイン カタログ]** を選択します。

1. Bing検索アドインの `\\ServerName\AppManifests` パスを既に追加している場合は、この手順をスキップします。 それ以外の場合は、[**信頼されたアドイン カタログ**] ウィンドウで、[**カタログの URL]** テキスト ボックスにパスを追加`\\ServerName\AppManifests`し、[**カタログの追加]** を選択し、既定のソースとしてネットワーク共有を有効にしてから (図 1 を参照)、[OK] を選択 **します**。

   *図 1.アドイン マニフェスト用のネットワーク ファイル共有の追加*

   ![アプリ マニフェストのネットワーク ファイル共有を追加する。](../images/pj15-create-simple-agave-manage-catalogs.png)

1. 新しいアドインを追加するか、ソース コードを変更したら、Project を再起動します。[**プロジェクト**] リボンで、[**Office アドイン**] ドロップダウン メニューの [**すべて表示**] を選択します。[**アドインの挿入**] ダイアログ ボックスで、[**共有フォルダー**] を選択し (図 2 を参照)、[**Project OM Test**]、[**挿入**] の順に選択します。Project OM Test アドインが作業ウィンドウ内で起動します。

   *図 2.ファイル共有上にある Project OM Test アドインの開始*

   ![アプリの挿入。](../images/pj15-create-simple-agave-start-agave-app.png)

1. Project で、少なくとも 2 つのタスクを備えた単純なプロジェクトを作成して保存します。 たとえば、T1 とT2 というタスク、およびM1 というマイルストーンを作成し、タスクの期間と先行タスクを図 3 のように設定します。 リボンの [**プロジェクト**] タブを選択し、タスク T2 の行全体を選択して、作業ウィンドウの [**getSelectedDataAsync**] ボタンを選択します。 図 3 に、 **Project OM Test** アドインのテキスト ボックス内で選択されているデータを示します。

   *図 3.Project OM Test アドインの使用*

   ![Project OM テスト アプリの使用。](../images/pj15-create-simple-agave-project-om-test.png)

1. 最初のタスクの [**期間**] 列内にあるセルを選択し、**Project OM Test** アドイン内の [**getSelectedDataAsync**] ボタンを選択します。 この関数は `getSelectedDataAsync` 、表示 `2 days`するテキスト ボックスの値を設定します。

1. 3 つのタスクすべての [**期間**] セル (3 つ) を選択します。 この関数は `getSelectedDataAsync` 、異なる行で選択されたセルに対してセミコロンで区切られたテキスト値を返します (例: `2 days;4 days;0 days`.

   この関数は `getSelectedDataAsync` 、行内で選択されたセルのコンマ区切りのテキスト値を返します。 たとえば、図 3 ではタスク T2 の行全体が選択されています。 選択 `getSelectedDataAsync`すると、テキスト ボックスに次の内容が表示されます。  `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`

   [**インジケーター**] 列と [**リソース名**] 列はどちらも空なので、テキスト配列にはこれらの列に対応する空の値が表示されます。 [`<NA>`] セルの値は [] です。

1. タスク T2 の行の任意のセル、またはタスク T2 の行全体を選択し、[**getSelectedTaskAsync**] を選択します。 テキスト ボックスにタスクの GUID 値が表示されます (例:  `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`)。 Project は、その値を **Project OM Test** アドインのグローバル`taskGuid`変数に格納します。

1. を選択します `getTaskAsync`。 変数に `taskGuid` タスク T2 の GUID が含まれている場合、テキスト ボックスにタスク情報が表示されます。 **ResourceNames** 値は空です。

    2 つのローカル リソース R1 と R2 を作成し、それぞれ 50% のタスク T2 に割り当て、 **getTaskAsync** をもう一度選択します。 テキスト ボックスの結果にはリソース情報が含まれます。 結果が同期された SharePoint タスク リスト内にある場合は、SharePoint のタスク ID も結果に含まれます。

    - タスク名: `T2`
    - GUID: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`
    - WSS Id: `0`
    - ResourceNames: `R1[50%],R2[50%]`

1. [ **タスク フィールドの取得]** ボタンを選択します。 この関数は `getTaskFields` 、タスク名、インデックス、開始日、期間、優先度、タスク ノートに対してメソッドを複数回呼び出 `getTaskFieldAsync` します。

    - 名前: `T2`
    - ID: `2`
    - 開始: `Thu 6/14/12`
    - 期間: `4d`
    - 優先度: `500`
    - ノート: これは、タスク T2 のノートです。 単なるテスト ノートです。 実際のノートの場合は、実際の情報になります。

1. **[getWSSUrlAsync]** ボタンを選択します。プロジェクトが次の種類のどちらかであれば、タスク リストの URL と名前が結果に表示されます。

    - Project Server にインポートされた SharePoint タスク リスト
    - Project Professional にインポートされ、SharePoint に (Project Server を使用せずに) 保存された SharePoint タスク リスト

    > [!NOTE]
    > Project Professional が Windows Server コンピューターにインストールされており、プロジェクトを SharePoint に保存できる場合は、**サーバー マネージャー** を使用して **デスクトップ エクスペリエンス** 機能を追加できます。

    プロジェクトがローカル プロジェクトの場合、またはProject Professionalを使用して Project Server によって管理されているプロジェクトを開く場合、`getWSSUrlAsync`メソッドは未定義のエラーを示します。

    - SharePoint URL: `http://ServerName`
    - リスト名: `Test task list`

1. **TaskSelectionChanged イベント** セクションで **[追加]** ボタンを選択します。これは、タスク選択変更イベントを登録する関数を呼び出`manageTaskEventHandler`し、テキスト ボックスに戻ります`In onComplete function for addHandlerAsync Status: succeeded`。 別のタスクを選択します。テキスト ボックスには、タスク選択変更イベントのコールバック関数の出力が表示 `In task selection changed event handler`されます。 [ **削除** ] ボタンを選択して、イベント ハンドラーの登録を解除します。

1. リソースに関するメソッドを使用するには、最初に [**リソース シート**]、[**リソース配分状況**]、[**リソース フォーム**] などのビューを選択し、次にそのビュー内でリソースを選択します。 **getSelectedResourceAsync** を選択して **resourceGuid** 変数を初期化し、[**リソース フィールドの取得**] を選択してリソース プロパティを複数回呼び出`getResourceFieldAsync`します。 また、リソース選択変更のイベント ハンドラーを追加または削除することもできます。

    - リソース名: `R1`
    - 原価: `$800.00`
    - 標準単価: `$50.00/h`
    - 実績コスト: `$0.00`
    - 実績作業時間 : `0h`
    - 単位: `100%`

1. **getSelectedViewAsync** を選択して、アクティブなビューの種類と名前を表示します。 また、ビュー選択変更のイベント ハンドラーを追加または削除することもできます。 たとえば、 **リソース フォーム** がアクティブなビューの場合、関数は `getSelectedViewAsync` テキスト ボックスに次の内容を表示します。

    - ビューの種類: `6`
    - 名前: `Resource Form`

1. [ **プロジェクト フィールドの取得** ] を選択して、作業中のプロジェクトのさまざまな `getProjectFieldAsync` プロパティに対して関数を複数回呼び出します。 プロジェクトがProject Web Appから開かれている場合、`getProjectFieldAsync`関数はProject Web App インスタンスの URL を取得できます。

    - プロジェクト GUID: `9845922E-DAB4-E111-8AF3-00155D3BA208`
    - 開始: `Tue 6/12/12`
    - 終了: `Tue 6/19/12`
    - 通貨桁数: `2`
    - 通貨記号: `$`
    - 記号の位置: `0`
    - Project Web App の URL: `http://servername/pwa`
  
1. **Office.Context.document** オブジェクトとオブジェクトのプロパティを取得して、ドキュメントとアドインが実行されているアプリケーションのプロパティを取得する [**コンテキスト値の取得**] ボタンを`Office.context.application`選択します。 For example, if the Project1.mpp file is on the local computer desktop, the document URL is `C:\Users\UserAlias\Desktop\Project1.mpp`. If the .mpp file is in a SharePoint library, the value is the URL of the document. If you use Project Professional 2013 to open a project named Project1 from Project Web App, the document URL is  `<>\Project1`.

    - ドキュメントの URL: `<>\Project1`
    - ドキュメント モード: `readWrite`
    - アプリの言語: `en-US`
    - 表示言語: `en-US`

1. ソース コードを編集した後は、Project をいったん閉じて再起動することで、アドインを最新の情報に更新できます。[**プロジェクト**] リボンの [**Office アドイン**] ドロップダウン リストに、最近使用したアドインの一覧が保持されています。

## <a name="example"></a>例

Project 2013 SDK のダウンロードには、JSOMCall.html ファイル、JSOM_Sample.js ファイル、関連する Office.js、Office.debug.js、Project-15.js、および Project-15.debug.js の各ファイルの完全なコードが含まれています。次に、JSOMCall.html ファイルのコードを示します。

```HTML
<!DOCTYPE html>
<html>
    <head>
        <title>Project OM Sample Code</title>
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

        <script type="text/javascript" src="MicrosoftAjax.js"></script>

        <!-- Use the CDN reference to office.js when deploying your add-in. -->
        <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
        <script type="text/javascript" src="office.js"></script>
        <script type="text/javascript" src="JSOM_Sample.js"></script>

        <style type="text/css">
            .button-wide {
                width: 210px;
                margin-top: 2px;
            }
            .button-narrow 
            {
                width: 80px;
                margin-top: 2px;
            }
        </style>
    </head>

    <body>
        <div id="Common_JSOM_API">
            OBJECT MODEL TESTS
            <br /><br />
            <strong>General method:</strong>
            <br />
            <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
                value="getSelectedDataAsync" />
        </div>
        <div id="ProjectSpecificTask">
            <br />
            <strong>Project-specific task methods:</strong><br />
            <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
            <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
            <strong>Task selection changed:</strong>
            <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ResourceMethods">
            <br />
            <strong>Resource methods:</strong>
            <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
            <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
            <strong>Resource selection changed:</strong>
            <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ViewMethods">
            <br />
            <strong>View method:</strong>
            <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
            <strong>View selection changed:</strong>
            <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ProjectMethods">
            <br />
            <strong>Project properties:</strong>
            <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
        </div>
        <div id="ContextVariables">
            <br />
            <strong>Context properties:</strong>
            <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
        </div>
        <br />
        <textarea id="text" rows="10" cols="25">This is the text result.</textarea>
    </body>
</html
```

## <a name="robust-programming"></a>堅牢なプログラミング

**Project OM テスト** アドインは、Project-15.js ファイルとOffice.js ファイルで Project 2013 用のいくつかの JavaScript 関数の使用を示す例です。 この例は単なるテスト用で、堅牢なエラー チェックは含まれていません。 たとえば、リソースを選択して関数を `getSelectedResourceAsync` 実行しない場合、 `resourceGuid` 変数は初期化されず、エラーを返す `getResourceFieldAsync` 呼び出しが行われます。 実際に運用するアドインでは、特定のエラーをチェックして結果を無視したり、特定の状況に該当しない機能を隠したり、機能を使用する前にビューや有効な項目を選択するようにユーザーに通知したりする必要があります。

簡単な例では、次のコードのエラー出力には、関数のエラーを回避するために実行するアクションを指定する th  `actionMessage` 変数が `getSelectedResourceAsync` 含まれています。

```js
function logError(errorText) {
    text.value = "Error in " + errorText;
}

function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);
}

// Get the GUID of the selected resource.
function getSelectedResourceAsync() {
    _projDoc.getSelectedResourceAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            resourceGuid = asyncResult.value;
        }
        else {
            var actionMessage = "Select a resource before running the getSelectedResourceAsync method.";
            logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                               asyncResult.error.message, actionMessage);
        }
    });
}
```

Project 2013 SDK のダウンロードの **HelloProject_OData** サンプルには、JQuery ライブラリを使用してポップアップ エラー メッセージを表示する SurfaceErrors.js ファイルが含まれています。 図 4 に、"toast" 通知のエラー メッセージを示します。

SurfaceErrors.js ファイルの次のコードには、オブジェクトを作成`Toast`する関数`throwError`が含まれています。

```js
/*
 * Show error messages in a "toast" notification.
 */

// Throws a custom defined error.
function throwError(errTitle, errMessage) {
    try {
        // Define and throw a custom error.
        var customError = { name: errTitle, message: errMessage }
        throw customError;
    }
    catch (err) {
        // Catch the error and display it to the user.
        Toast.showToast(err.name, err.message);
    }
}

// Add a dynamically-created div "toast" for displaying errors to the user.
var Toast = {

    Toast: "divToast",
    Close: "btnClose",
    Notice: "lblNotice",
    Output: "lblOutput",

    // Show the toast with the specified information.
    showToast: function (title, message) {

        if (document.getElementById(this.Toast) == null) {
            this.createToast();
        }

        document.getElementById(this.Notice).innerText = title;
        document.getElementById(this.Output).innerText = message;

        $("#" + this.Toast).hide();
        $("#" + this.Toast).show("slow");
    },

    // Create the display for the toast.
    createToast: function () {
        var divToast;
        var lblClose;
        var btnClose;
        var divOutput;
        var lblOutput;
        var lblNotice;

        // Create the container div.
        divToast = document.createElement("div");
        var toastStyle = "background-color:rgba(220, 220, 128, 0.80);" +
            "position:absolute;" +
            "bottom:0px;" +
            "width:90%;" +
            "text-align:center;" +
            "font-size:11pt;";
        divToast.setAttribute("style", toastStyle);
        divToast.setAttribute("id", this.Toast);

        // Create the close button.
        lblClose = document.createElement("div");
        lblClose.setAttribute("id", this.Close);
        var btnStyle = "text-align:right;" +
            "padding-right:10px;" +
            "font-size:10pt;" +
            "cursor:default";
        lblClose.setAttribute("style", btnStyle);
        lblClose.appendChild(document.createTextNode("CLOSE "));

        btnClose = document.createElement("span");
        btnClose.setAttribute("style", "cursor:pointer;");
        btnClose.setAttribute("onclick", "Toast.close()");
        btnClose.innerText = "X";
        lblClose.appendChild(btnClose);

        // Create the div to contain the toast title and message.
        divOutput = document.createElement("div");
        divOutput.setAttribute("id", "divOutput");
        var outputStyle = "margin-top:0px;";
        divOutput.setAttribute("style", outputStyle);

        lblNotice = document.createElement("span");
        lblNotice.setAttribute("id", this.Notice);
        var labelStyle = "font-weight:bold;margin-top:0px;";
        lblNotice.setAttribute("style", labelStyle);

        lblOutput = document.createElement("span");
        lblOutput.setAttribute("id", this.Output);

        // Add the child nodes to the toast div.
        divOutput.appendChild(lblNotice);
        divOutput.appendChild(document.createElement("br"));
        divOutput.appendChild(lblOutput);
        divToast.appendChild(lblClose);
        divToast.appendChild(divOutput);

        // Add the toast div to the document body.
        document.body.appendChild(divToast);
    },

    // Close the toast.
    close: function () {
        $("#" + this.Toast).hide("slow");
    }
}
```

この関数を`throwError`使用するには、JQuery ライブラリとSurfaceErrors.js スクリプトをJSOMCall.html ファイルに含め、次に`throwError`他の JavaScript 関数 (.`logMethodError`

> [!NOTE]
> アドインを展開する前に、office.js の参照と jQuery の参照をコンテンツ配信ネットワーク (CDN) の参照に変更してください。CDN の参照は最新のバージョンと高いパフォーマンスを提供します。

```HTML
<!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to Office.js and jQuery when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
    <script type="text/javascript" src="office.js"></script>
    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jQuery/jquery-1.9.0.min.js"></script>

    <script type="text/javascript" src="JSOM_Sample.js"></script>
    <script type="text/javascript" src="SurfaceErrors.js"></script>

    <!-- . . . INVALID USE OF SYMBOLS . . . -->
</head>
```

```js
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```

*図 4. SurfaceErrors.js ファイル内の関数は "toast" 通知を表示できます*

![SurfaceError ルーチンを使用してエラーを表示する。](../images/pj15-create-simple-agave-surface-error.png)

## <a name="see-also"></a>関連項目

- [Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)
- [アドイン用の JavaScript API について](../develop/understanding-the-javascript-api-for-office.md)
- [Office JavaScript API アドイン](../reference/javascript-api-for-office.md)
- [Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
- [Project 2013 SDK のダウンロード](https://www.microsoft.com/download/details.aspx?id=30435)
