---
title: テキスト エディターを使用して Microsoft Project 用の作業ウィンドウ アドインを初めて作成する
description: Project Standard アドイン用の Yeo Office man ジェネレーターを使用して、Project Professional 2013 以降のバージョンの作業ウィンドウ アドインを作成します。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: c1de70bec62c4080306c985a319601c506270f2b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348420"
---
# <a name="create-your-first-task-pane-add-in-for-microsoft-project-by-using-a-text-editor"></a><span data-ttu-id="3dfc0-103">テキスト エディターを使用して Microsoft Project 用の作業ウィンドウ アドインを初めて作成する</span><span class="sxs-lookup"><span data-stu-id="3dfc0-103">Create your first task pane add-in for Microsoft Project by using a text editor</span></span>

<span data-ttu-id="3dfc0-104">Office アドインの Yeoman ジェネレーターを使用して、Project Standard 2013、Project Professional 2013、または以降のバージョンの作業ウィンドウ アドインを作成できます。この記事では、ファイル共有上の HTML ファイルをポイントする XML マニフェストを使用する単純なアドインを作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-104">You can create a task pane add-in for Project Standard 2013, Project Professional 2013, or later versions using the Yeoman generator for Office Add-ins. This article describes how to create a simple add-in that uses an XML manifest that points to an HTML file on a file share.</span></span> <span data-ttu-id="3dfc0-105">OM Test Projectサンプル アドインは、アドインにオブジェクト モデルを使用する JavaScript 関数をテストします。Project の信頼センターを使用してマニフェスト ファイルを含むファイル共有を登録した後、リボンの [Project] タブから作業ウィンドウ **アドインを** 開きます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-105">The Project OM Test sample add-in tests some JavaScript functions that use the object model for add-ins. After you use the **Trust Center** in Project to register the file share that contains the manifest file, you can open the task pane add-in from the **Project** tab on the ribbon.</span></span> <span data-ttu-id="3dfc0-106">Project 2013の セキュリティ センターを使用して、マニフェスト ファイルを含むファイル共有を登録した後は、作業ウィンドウ アドインをリボンの [ プロジェクト] タブから開くことができます (この記事のサンプル コードは、Microsoft Corporation の Arvind Iyer によるテスト アプリケーションに基づくものです)。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-106">(The sample code in this article is based on a test application by Arvind Iyer, Microsoft Corporation.)</span></span>

<span data-ttu-id="3dfc0-107">Projectは、他のクライアントが使用するのと同じアドイン マニフェスト スキーマOffice JavaScript API の多くを使用します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-107">Project uses the same add-in manifest schema that other Office clients use, and much of the same JavaScript API.</span></span> <span data-ttu-id="3dfc0-108">この記事に記載されているアドインの完全なコードは、Project 2013 SDK ダウンロードのサブディレクトリ `Samples\Apps` で提供されています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-108">The complete code for the add-in that is described in this article is available in the  `Samples\Apps` subdirectory of the Project 2013 SDK download.</span></span>

<span data-ttu-id="3dfc0-109">Project OM Test サンプル アドインは、タスクの GUID と、アプリケーションおよびアクティブなプロジェクトのプロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-109">The Project OM Test sample add-in can get the GUID of a task and properties of the application and the active project.</span></span> <span data-ttu-id="3dfc0-110">Project Professional 2013 で SharePoint ライブラリ内にあるプロジェクトを開くと、このアドインでは、そのプロジェクトの URL を表示できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-110">If Project Professional 2013 opens a project that is in a SharePoint library, the add-in can show the URL of the project.</span></span> 

<span data-ttu-id="3dfc0-p104">[Project 2013 SDK のダウンロード](https://www.microsoft.com/download/details.aspx?id=30435%20)には完全なソース コードが含まれています。Project2013SDK.msi に含まれる SDK を展開してインストールしたら、`\Samples\Apps\Copy_to_AppManifests_FileShare` サブディレクトリにマニフェスト ファイルがあり、`\Samples\Apps\Copy_to_AppSource_FileShare` サブディレクトリにソース コードがあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p104">The [Project 2013 SDK download](https://www.microsoft.com/download/details.aspx?id=30435%20) includes the complete source code. When you extract and install the SDK and samples that are in the Project2013SDK.msi file, see the `\Samples\Apps\Copy_to_AppManifests_FileShare` subdirectory for the manifest file and the `\Samples\Apps\Copy_to_AppSource_FileShare` subdirectory for the source code.</span></span> 

<span data-ttu-id="3dfc0-113">サンプルの JSOMCall.html では、インクルードされる office.js ファイルと project-15.js ファイル内の JavaScript 関数を使用しています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-113">The JSOMCall.html sample uses JavaScript functions in the office.js file and project-15.js file, which are included.</span></span> <span data-ttu-id="3dfc0-114">対応するデバッグ ファイル (office.debug.js および project-15.debug.js) を使用すると、これらの関数を検証できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-114">You can use the corresponding debug files (office.debug.js and project-15.debug.js) to examine the functions.</span></span>

<span data-ttu-id="3dfc0-115">アドインでの JavaScript の使用の概要Office JavaScript API の概要[Office参照してください](../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-115">For an introduction to using JavaScript in Office Add-ins, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a><span data-ttu-id="3dfc0-p106">手順 1. アドイン マニフェスト ファイルを作成するには</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p106">Procedure 1. To create the add-in manifest file</span></span>

<span data-ttu-id="3dfc0-118">ローカル ディレクトリに XML ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-118">Create an XML file in a local directory.</span></span> <span data-ttu-id="3dfc0-119">XML ファイルには、要素要素と子要素が含まれます。これは、アドイン XML マニフェストのOffice `OfficeApp` [で説明されています](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-119">The XML file includes the `OfficeApp` element and child elements, which are described in the [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span> <span data-ttu-id="3dfc0-120">たとえば、次の XML を含む JSOM_SimpleOMCalls.xmlという名前のファイルを作成します (要素の GUID 値を変更 `Id` します)。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-120">For example, create a file named JSOM_SimpleOMCalls.xml that contains the following XML (change the GUID value of the `Id` element).</span></span>

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

<span data-ttu-id="3dfc0-121">このProject要素 `OfficeApp` に属性値を含める `xsi:type="TaskPaneApp"` 必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-121">For Project, the `OfficeApp` element must include the `xsi:type="TaskPaneApp"` attribute value.</span></span> <span data-ttu-id="3dfc0-122">要素 `Id` は GUID です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-122">The `Id` element is a GUID.</span></span> <span data-ttu-id="3dfc0-123">この値は、アドイン HTML ソース ファイルSharePoint作業ウィンドウで実行される Web アプリケーションのファイル共有パスまたは URL である `SourceLocation` 必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-123">The `SourceLocation` value must be a file share path or a SharePoint URL for the add-in HTML source file or the web application that runs in the task pane.</span></span> <span data-ttu-id="3dfc0-124">For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="3dfc0-124">For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).</span></span>

<span data-ttu-id="3dfc0-p109">手順 2. では、JSOM_SimpleOMCalls.xml マニフェストが Project テスト アドインのために指定する HTML ファイルの作成方法を示します。この HTML 内で指定されているボタンは、関連する JavaScript 関数を呼び出します。JavaScript 関数は、この HTML ファイル内に追加したり、別の .js ファイル内に配置したりできます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p109">Procedure 2 shows how to create the HTML file that the JSOM_SimpleOMCalls.xml manifest specifies for the Project test add-in. Buttons that are specified in the HTML file call related JavaScript functions. You can add the JavaScript functions within the HTML file, or put them in a separate .js file.</span></span>

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a><span data-ttu-id="3dfc0-p110">手順 2. Project OM Test アドインのソース ファイルを作成するには</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p110">Procedure 2. To create the source files for the Project OM Test add-in</span></span>

1. <span data-ttu-id="3dfc0-130">マニフェスト内の要素で指定された名前の HTML `SourceLocation` ファイルをJSOM_SimpleOMCalls.xmlします。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-130">Create an HTML file with a name that is specified by the `SourceLocation` element in the JSOM_SimpleOMCalls.xml manifest.</span></span>

   <span data-ttu-id="3dfc0-131">たとえば、`C:\Project\AppSource`ディレクトリで theJSOMCall.html ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-131">For example, create theJSOMCall.html file in the `C:\Project\AppSource` directory.</span></span> <span data-ttu-id="3dfc0-132">単純なテキスト エディターを使用してソース ファイルを作成することもできますが、特定の種類のドキュメント (HTML や JavaScript など) で動作し、その他の編集補助機能を備え、Visual Studio Code などのツールを使用する方が簡単です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-132">Although you can use a simple text editor to create the source files, it is easier to use a tool such as Visual Studio Code, which works with specific document types (such as HTML and JavaScript) and has other editing aids.</span></span> <span data-ttu-id="3dfc0-133">「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」で説明されている Bing Search の例をまだ行っていない場合は、マニフェストが指定する `\\ServerName\AppSource` ファイル共有を作成する方法が手順 3 で示されています。　</span><span class="sxs-lookup"><span data-stu-id="3dfc0-133">If you have not already done the Bing Search example that is described in [Task pane add-ins for Project](../project/project-add-ins.md), Procedure 3 shows how to create the `\\ServerName\AppSource` file share that the manifest specifies.</span></span>

   <span data-ttu-id="3dfc0-134">JSOMCall.html ファイルは、AJAX 機能に共通の MicrosoftAjax.js ファイルを使用し、Office.js ファイルを 2013 アプリケーションのアドイン機能に使用Officeします。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-134">The JSOMCall.html file uses the common MicrosoftAjax.js file for AJAX functionality and the Office.js file for the add-in functionality in Office 2013 applications.</span></span>

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

   <span data-ttu-id="3dfc0-135">要素 `textarea` は、JavaScript 関数の結果を示すテキスト ボックスを指定します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-135">The `textarea` element specifies a text box that shows results of the JavaScript functions.</span></span>

   > [!NOTE]
   > <span data-ttu-id="3dfc0-136">Project OM Test サンプルを実行するには、Project 2013 SDK ダウンロードに含まれるファイル (Office.js、Project-15.js、および MicrosoftAjax.js) を JSOMCall.html ファイルと同じディレクトリにコピーします。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-136">For the Project OM Test sample to work, copy the following files from the Project 2013 SDK download to the same directory as the JSOMCall.html file: Office.js, Project-15.js, and MicrosoftAjax.js.</span></span>

   <span data-ttu-id="3dfc0-p112">手順 2. では、Project OM Test サンプル アドインが使用する特定の関数のために JSOM Sample.js というファイルを追加します。この後の手順では、JavaScript 関数を呼び出すボタン用にその他の HTML 要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p112">Step 2 adds the JSOM_Sample.js file for specific functions that the Project OM Test sample add-in uses. In later steps, you will add other HTML elements for buttons that call JavaScript functions.</span></span>

1. <span data-ttu-id="3dfc0-139">JSOM_Sample.js という名前の JavaScript ファイルを、JSOMCall.html ファイルと同じディレクトリ内に作成します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-139">Create a JavaScript file named JSOM_Sample.js in the same directory as the JSOMCall.html file.</span></span>

   <span data-ttu-id="3dfc0-140">次のコードは、Office.js ファイル内の関数を使用して、アプリケーションのコンテキストとドキュメント情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-140">The following code gets the application context and document information by using functions in the Office.js file.</span></span> <span data-ttu-id="3dfc0-141">オブジェクト `text` は、HTML ファイル `textarea` 内のコントロールの ID です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-141">The `text` object is the ID of the `textarea` control in the HTML file.</span></span>

   <span data-ttu-id="3dfc0-142">**\_ projDoc 変数** はオブジェクトで初期化 `ProjectDocument` されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-142">The **\_projDoc** variable is initialized with a `ProjectDocument` object.</span></span> <span data-ttu-id="3dfc0-143">このコードには、いくつかの単純なエラー処理関数と、アプリケーション コンテキストとプロジェクト ドキュメント コンテキスト プロパティを取得 `getContextValues` する関数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-143">The code includes some simple error handling functions, and the `getContextValues` function that gets application context and project document context properties.</span></span> <span data-ttu-id="3dfc0-144">Project の JavaScript オブジェクト モデルの詳細については、「[JavaScript API for Office](../reference/javascript-api-for-office.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-144">For more information about the JavaScript object model for Project, see [JavaScript API for Office](../reference/javascript-api-for-office.md).</span></span>


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
        // Checks for the DOM to load using the jQuery ready function.
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

   <span data-ttu-id="3dfc0-145">ファイル内の関数の詳細については、「Office.debug.js [JavaScript API Office参照してください](../reference/javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-145">For information about the functions in the Office.debug.js file, see [Office JavaScript API](../reference/javascript-api-for-office.md).</span></span> <span data-ttu-id="3dfc0-146">たとえば、関数は `getDocumentUrl` 開いているプロジェクトの URL またはファイル パスを取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-146">For example, the `getDocumentUrl` function gets the URL or file path of the open project.</span></span>

1. <span data-ttu-id="3dfc0-147">Office.js および Project-15.js 内の非同期関数を呼び出して選択されているデータを取得する JavaScript 関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-147">Add JavaScript functions that call asynchronous functions in Office.js and Project-15.js to get selected data:</span></span>

   - <span data-ttu-id="3dfc0-148">たとえば、選択したデータの書式設定されていないOffice.jsを取得する関数の一 `getSelectedDataAsync` 般的な関数です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-148">For example, `getSelectedDataAsync` is a general function in Office.js that gets unformatted text for the selected data.</span></span> <span data-ttu-id="3dfc0-149">詳細については、[「AsyncResult オブジェクト」](/javascript/api/office/office.asyncresult)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-149">For more information, see [AsyncResult object](/javascript/api/office/office.asyncresult).</span></span>

   - <span data-ttu-id="3dfc0-150">この `getSelectedTaskAsync` 関数はProject-15.jsタスクの GUID を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-150">The `getSelectedTaskAsync` function in Project-15.js gets the GUID of the selected task.</span></span> <span data-ttu-id="3dfc0-151">同様に、 `getSelectedResourceAsync` 関数は選択したリソースの GUID を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-151">Similarly, the `getSelectedResourceAsync` function gets the GUID of the selected resource.</span></span> <span data-ttu-id="3dfc0-152">タスクまたはリソースが選択されていない状態でこれらの関数を呼び出すと、未定義のエラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-152">If you call those functions when a task or a resource is not selected, the functions show an undefined error.</span></span>

   - <span data-ttu-id="3dfc0-153">関数 `getTaskAsync` は、タスク名と割り当てられたリソースの名前を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-153">The `getTaskAsync` function gets the task name and the names of the assigned resources.</span></span> <span data-ttu-id="3dfc0-154">タスクが同期されたタスク リストにあるSharePoint、SharePoint リスト内のタスク ID を取得します。それ以外の場合、タスク ID は 0 になりますSharePoint `getTaskAsync` します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-154">If the task is in a synchronized SharePoint task list, `getTaskAsync` gets the task ID in the SharePoint list; otherwise, the SharePoint task ID is 0.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3dfc0-155">サンプル コードには、デモ用にバグが含まれています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-155">For demonstration purposes, the example code includes a bug.</span></span> <span data-ttu-id="3dfc0-156">未定義 `taskGuid` の場合は、 `getTaskAsync` 関数エラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-156">If `taskGuid` is undefined, the `getTaskAsync` function errors off.</span></span> <span data-ttu-id="3dfc0-157">有効なタスク GUID を取得し、別のタスクを選択すると、関数によって操作された最新のタスクのデータ `getTaskAsync` が取得 `getSelectedTaskAsync` されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-157">If you get a valid task GUID and then select a different task, the `getTaskAsync` function gets data for the most recent task that was operated on by the `getSelectedTaskAsync` function.</span></span>
  
   - <span data-ttu-id="3dfc0-158">`getTaskFields`、、およびタスクまたはリソースの指定されたフィールドを取得する、または複数回呼び出す `getResourceFields` `getProjectFields` `getTaskFieldAsync` `getResourceFieldAsync` `getProjectFieldAsync` ローカル関数です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-158">`getTaskFields`, `getResourceFields`, and `getProjectFields` are local functions that call `getTaskFieldAsync`, `getResourceFieldAsync`, or `getProjectFieldAsync` multiple times to get specified fields of a task or a resource.</span></span> <span data-ttu-id="3dfc0-159">このファイルproject-15.debug.js、列挙 `ProjectTaskFields` 体と列挙には、サポートされている `ProjectResourceFields` フィールドが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-159">In the project-15.debug.js file, the `ProjectTaskFields` enumeration and the `ProjectResourceFields` enumeration show which fields are supported.</span></span>

   - <span data-ttu-id="3dfc0-160">この `getSelectedViewAsync` 関数は、ビューの種類 (project-15.debug.js の列挙で定義 `ProjectViewTypes` ) とビューの名前を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-160">The `getSelectedViewAsync` function gets the type of view (defined in the `ProjectViewTypes` enumeration in project-15.debug.js) and the name of the view.</span></span>

   - <span data-ttu-id="3dfc0-161">プロジェクトがタスク リストと同期SharePoint、関数は URL とタスク リスト `getWSSUrlAsync` の名前を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-161">If the project is synchronized with a SharePoint tasks list, the `getWSSUrlAsync` function gets the URL and the name of the tasks list.</span></span> <span data-ttu-id="3dfc0-162">プロジェクトがタスク リストと同期されていない場合、SharePointエラー `getWSSUrlAsync` が発生します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-162">If the project is not synchronized with a SharePoint tasks list, the `getWSSUrlAsync` function errors off.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3dfc0-163">タスクリストのSharePoint URL と名前を取得するには `getProjectFieldAsync` `WSSUrl` `WSSList` [、ProjectProjectFields](/javascript/api/office/office.projectprojectfields)列挙の and 定数と一緒に関数を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-163">To get the SharePoint URL and name of the tasks list, we recommend that you use the `getProjectFieldAsync` function with the `WSSUrl` and `WSSList` constants in the [ProjectProjectFields](/javascript/api/office/office.projectprojectfields) enumeration.</span></span>

   <span data-ttu-id="3dfc0-p122">次のコードの各関数には、`function (asyncResult)` によって指定されている匿名関数が含まれます。これは、非同期の結果を取得するコールバックです。匿名関数の代わりに、複雑なアドインの保守に役立つ名前付き関数を使用できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p122">Each of the functions in the following code includes an anonymous function that is specified by  `function (asyncResult)`, which is a callback that gets the asynchronous result. Instead of anonymous functions, you could use named functions, which can help with maintainability of complex add-ins.</span></span>

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

1. <span data-ttu-id="3dfc0-166">JavaScript イベント ハンドラー コールバックおよび関数を追加して、タスク選択、リソース選択、およびビュー選択の変更に関するイベント ハンドラーの登録と登録解除を行います。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-166">Add JavaScript event handler callbacks and functions to register the task selection, resource selection, and view selection change event handlers and to unregister the event handlers.</span></span> <span data-ttu-id="3dfc0-167">この `manageEventHandlerAsync` 関数は、operation パラメーターに応じて、指定したイベント ハンドラーを追加または _削除_ します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-167">The `manageEventHandlerAsync` function adds or removes the specified event handler, depending on the _operation_ parameter.</span></span> <span data-ttu-id="3dfc0-168">操作は、 または `addHandlerAsync` `removeHandlerAsync` です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-168">The operation can be `addHandlerAsync` or `removeHandlerAsync`.</span></span>

   <span data-ttu-id="3dfc0-169">、 `manageTaskEventHandler` `manageResourceEventHandler` 、および `manageViewEventHandler` 関数は _、docMethod_ パラメーターで指定されたイベント ハンドラーを追加または削除できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-169">The `manageTaskEventHandler`, `manageResourceEventHandler`, and `manageViewEventHandler` functions can add or remove an event handler, as specified by the _docMethod_ parameter.</span></span>

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

1. <span data-ttu-id="3dfc0-170">この HTML ドキュメントの本文に、テストのために JavaScript 関数を呼び出すボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-170">For the body of the HTML document, add buttons that call the JavaScript functions for testing.</span></span> <span data-ttu-id="3dfc0-171">たとえば、共通 JSOM API の要素に、汎用関数を呼び出す `div` 入力ボタンを追加 `getSelectedDataAsync` します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-171">For example, in the `div` element for the common JSOM API, add an input button that calls the general `getSelectedDataAsync` function.</span></span>

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

1. <span data-ttu-id="3dfc0-172">プロジェクト固有 `div` のタスク関数とイベントのボタンを含むセクションを追加 `TaskSelectionChanged` します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-172">Add a `div` section with buttons for project-specific task functions and for the `TaskSelectionChanged` event.</span></span>

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

1. <span data-ttu-id="3dfc0-173">リソース メソッドとイベント、ビュー メソッドとイベント、プロジェクト プロパティ、コンテキスト プロパティのボタンを含むセクション `div` を追加する</span><span class="sxs-lookup"><span data-stu-id="3dfc0-173">Add `div` sections with buttons for the resource methods and events, view methods and events, project properties, and context properties</span></span>

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

1. <span data-ttu-id="3dfc0-174">ボタン要素の書式を設定するには、CSS 要素を追加 `style` します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-174">To format the button elements, add a CSS `style` element.</span></span> <span data-ttu-id="3dfc0-175">たとえば、要素の子として次のように追加 `head` します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-175">For example, add the following as a child of the `head` element.</span></span>

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

<span data-ttu-id="3dfc0-176">手順 3. では、Project OM Test アドインの機能をインストールして使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-176">Procedure 3 shows how to install and use the Project OM Test add-in features.</span></span>

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a><span data-ttu-id="3dfc0-p126">手順 3. Project OM Test アドインをインストールして使用するには</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p126">Procedure 3. To install and use the Project OM Test add-in</span></span>

1. <span data-ttu-id="3dfc0-179">JSOM SimpleOMCalls.xml マニフェストが含まれているディレクトリに対するファイル共有を作成します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-179">Create a file share for the directory that contains the JSOM_SimpleOMCalls.xml manifest.</span></span> <span data-ttu-id="3dfc0-180">ファイル共有は、ローカル コンピューター上、またはネットワーク上のアクセス可能なリモート コンピューター上に作成できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-180">You can create the file share on the local computer or on a remote computer that is accessible on the network.</span></span> <span data-ttu-id="3dfc0-181">たとえば、マニフェストがローカル コンピューターのディレクトリにある場合は、  `C:\Project\AppManifests` 次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-181">For example, if the manifest is in the  `C:\Project\AppManifests` directory on the local computer, run the following command.</span></span>

    `Net share AppManifests=C:\Project\AppManifests`

1. <span data-ttu-id="3dfc0-182">Project OM Test アドインの HTML および JavaScript ファイルが含まれるディレクトリに対するファイル共有を作成します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-182">Create a file share for the directory that contains the HTML and JavaScript files for the Project OM Test add-in.</span></span> <span data-ttu-id="3dfc0-183">このファイル共有パスは、JSOM SimpleOMCalls.xml マニフェストで指定されているパスに一致するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-183">Ensure the file share path matches the path that is specified in the JSOM_SimpleOMCalls.xml manifest.</span></span> <span data-ttu-id="3dfc0-184">たとえば、ファイルがローカル コンピューターのディレクトリにある場合は、  `C:\Project\AppSource` 次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-184">For example, if the files are in the  `C:\Project\AppSource` directory on the local computer, run the following command.</span></span>

    `net share AppSource=C:\Project\AppSource`

1. <span data-ttu-id="3dfc0-185">Project で、[**Project のオプション**] ダイアログ ボックスを開き、[**セキュリティ センター**]、[**セキュリティ センターの設定**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-185">In Project, open the **Project Options** dialog box, choose **Trust Center**, and then choose **Trust Center Settings**.</span></span>

   <span data-ttu-id="3dfc0-186">アドインの登録手順および追加情報については、「[Project 用の作業ウィンドウ アドイン](../project/project-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-186">The procedure for registering an add-in is also described in [Task pane add-ins for Project](../project/project-add-ins.md), with additional information.</span></span>

1. <span data-ttu-id="3dfc0-187">**[セキュリティ センター]** ダイアログ ボックスの左側のウィンドウで、**[信頼されているアドイン カタログ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-187">In the **Trust Center** dialog box, in the left pane, choose **Trusted Add-in Catalogs**.</span></span>

1. <span data-ttu-id="3dfc0-188">検索アドインのパスを既に追加しているBing、この手順 `\\ServerName\AppManifests` をスキップします。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-188">If you have already added the `\\ServerName\AppManifests` path for the Bing Search add-in, skip this step.</span></span> <span data-ttu-id="3dfc0-189">それ以外の場合は、[信頼できるアドイン カタログ] ウィンドウで、[カタログ URL] テキスト ボックスにパスを追加し、[カタログの追加] を選択し、ネットワーク共有を既定のソースとして有効にします (図 1 を参照 `\\ServerName\AppManifests` **)、[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-189">Otherwise, in the **Trusted Add-in Catalogs** pane, add the `\\ServerName\AppManifests` path in the **Catalog Url** text box, choose **Add catalog**, enable the network share as a default source (see Figure 1), and then choose **OK**.</span></span>

   <span data-ttu-id="3dfc0-190">*図 1.アドイン マニフェスト用のネットワーク ファイル共有の追加*</span><span class="sxs-lookup"><span data-stu-id="3dfc0-190">*Figure 1. Adding a network file share for add-in manifests*</span></span>

   ![アプリ マニフェストのネットワーク ファイル共有を追加する。](../images/pj15-create-simple-agave-manage-catalogs.png)

1. <span data-ttu-id="3dfc0-p130">新しいアドインを追加するか、ソース コードを変更したら、Project を再起動します。[**プロジェクト**] リボンで、[**Office アドイン**] ドロップダウン メニューの [**すべて表示**] を選択します。[**アドインの挿入**] ダイアログ ボックスで、[**共有フォルダー**] を選択し (図 2 を参照)、[**Project OM Test**]、[**挿入**] の順に選択します。Project OM Test アドインが作業ウィンドウ内で起動します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p130">After you add new add-ins, or change the source code, restart Project. On the **PROJECT** ribbon, choose the **Office Add-ins** drop-down menu, and then choose **See All**. In the **Insert Add-in** dialog box, choose **SHARED FOLDER** (see Figure 2), select **Project OM Test**, and then choose **Insert**. The Project OM Test add-in starts in a task pane.</span></span>

   <span data-ttu-id="3dfc0-196">*図 2.ファイル共有上にある Project OM Test アドインの開始*</span><span class="sxs-lookup"><span data-stu-id="3dfc0-196">*Figure 2. Starting the Project OM Test add-in that is on a file share*</span></span>

   ![アプリの挿入。](../images/pj15-create-simple-agave-start-agave-app.png)

1. <span data-ttu-id="3dfc0-198">Project で、少なくとも 2 つのタスクを備えた単純なプロジェクトを作成して保存します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-198">In Project, create and save a simple project that has at least two tasks.</span></span> <span data-ttu-id="3dfc0-199">たとえば、T1 とT2 というタスク、およびM1 というマイルストーンを作成し、タスクの期間と先行タスクを図 3 のように設定します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-199">For example, create tasks named T1, T2, and a milestone named M1, and then set the task durations and predecessors to be similar to those in Figure 3.</span></span> <span data-ttu-id="3dfc0-200">リボンの [**プロジェクト**] タブを選択し、タスク T2 の行全体を選択して、作業ウィンドウの [**getSelectedDataAsync**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-200">Choose the **PROJECT** tab on the ribbon, select the entire row for task T2, and then choose the **getSelectedDataAsync** button in the task pane.</span></span> <span data-ttu-id="3dfc0-201">図 3 に、 **Project OM Test** アドインのテキスト ボックス内で選択されているデータを示します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-201">Figure 3 shows the data that is selected in the text box of the **Project OM Test** add-in.</span></span>

   <span data-ttu-id="3dfc0-202">*図 3.Project OM Test アドインの使用*</span><span class="sxs-lookup"><span data-stu-id="3dfc0-202">*Figure 3. Using the Project OM Test add-in*</span></span>

   ![OM テスト アプリProject使用します。](../images/pj15-create-simple-agave-project-om-test.png)

1. <span data-ttu-id="3dfc0-204">最初のタスクの [**期間**] 列内にあるセルを選択し、**Project OM Test** アドイン内の [**getSelectedDataAsync**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-204">Select the cell in the **Duration** column for the first task, and then choose the **getSelectedDataAsync** button in the **Project OM Test** add-in.</span></span> <span data-ttu-id="3dfc0-205">この `getSelectedDataAsync` 関数は、テキスト ボックスの値を表示に設定します `2 days` 。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-205">The `getSelectedDataAsync` function sets the text box value to show `2 days`.</span></span> 

1. <span data-ttu-id="3dfc0-206">3 つのタスクすべての [**期間**] セル (3 つ) を選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-206">Select the three **Duration** cells for all three tasks.</span></span> <span data-ttu-id="3dfc0-207">この関数は、異なる行で選択されたセルのセミコロンで区切られたテキスト値を `getSelectedDataAsync` 返します `2 days;4 days;0 days` 。たとえば、 。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-207">The `getSelectedDataAsync` function returns semicolon-separated text values for cells selected in different rows, for example, `2 days;4 days;0 days`.</span></span>

   <span data-ttu-id="3dfc0-208">この `getSelectedDataAsync` 関数は、行内で選択されたセルのコンマ区切りテキスト値を返します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-208">The `getSelectedDataAsync` function returns comma-separated text values for cells selected within a row.</span></span> <span data-ttu-id="3dfc0-209">たとえば、図 3 ではタスク T2 の行全体が選択されています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-209">For example in Figure 3, the entire row for task T2 is selected.</span></span> <span data-ttu-id="3dfc0-210">選択すると、 `getSelectedDataAsync` 次のテキスト ボックスが表示されます。  `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-210">When you choose `getSelectedDataAsync`, the text box shows the following:  `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`</span></span>

   <span data-ttu-id="3dfc0-211">[**インジケーター**] 列と [**リソース名**] 列はどちらも空なので、テキスト配列にはこれらの列に対応する空の値が表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-211">The **Indicators** column and the **Resource Names** column are both empty, so the text array shows empty values for those columns.</span></span> <span data-ttu-id="3dfc0-212">[`<NA>`] セルの値は [] です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-212">The `<NA>` value is for the **Add New Column** cell.</span></span>

1. <span data-ttu-id="3dfc0-213">タスク T2 の行の任意のセル、またはタスク T2 の行全体を選択し、[**getSelectedTaskAsync**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-213">Select any cell in the row for task T2, or the entire row for task T2, and then choose **getSelectedTaskAsync**.</span></span> <span data-ttu-id="3dfc0-214">テキスト ボックスにタスクの GUID 値が表示されます (例:  `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`)。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-214">The text box shows the task GUID value, for example,  `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`.</span></span> <span data-ttu-id="3dfc0-215">Project OM Test アドインのグローバル変数に値Project `taskGuid` **格納** します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-215">Project stores that value in the global `taskGuid` variable of the **Project OM Test** add-in.</span></span>

1. <span data-ttu-id="3dfc0-216">を選択します `getTaskAsync` 。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-216">Select `getTaskAsync`.</span></span> <span data-ttu-id="3dfc0-217">変数にタスク T2 の GUID が含まれている場合、 `taskGuid` テキスト ボックスにタスク情報が表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-217">If the `taskGuid` variable contains the GUID for task T2, the text box displays the task information.</span></span> <span data-ttu-id="3dfc0-218">**ResourceNames** 値は空です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-218">The **ResourceNames** value is empty.</span></span>

    <span data-ttu-id="3dfc0-219">2 つのローカル リソース R1 と R2 を作成し、それぞれ 50% でタスク T2 に割り当て、 **再度 getTaskAsync を選択** します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-219">Create two local resources R1 andR2, assign them to task T2 at 50% each, and choose **getTaskAsync** again.</span></span> <span data-ttu-id="3dfc0-220">テキスト ボックスの結果にはリソース情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-220">The results in the text box include the resource information.</span></span> <span data-ttu-id="3dfc0-221">結果が同期された SharePoint タスク リスト内にある場合は、SharePoint のタスク ID も結果に含まれます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-221">If the task is in a synchronized SharePoint task list, the results also include the SharePoint task ID.</span></span>

    - <span data-ttu-id="3dfc0-222">タスク名: `T2`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-222">Task name: `T2`</span></span>
    - <span data-ttu-id="3dfc0-223">GUID: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-223">GUID: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`</span></span>
    - <span data-ttu-id="3dfc0-224">WSS Id: `0`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-224">WSS Id: `0`</span></span>
    - <span data-ttu-id="3dfc0-225">ResourceNames: `R1[50%],R2[50%]`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-225">ResourceNames: `R1[50%],R2[50%]`</span></span>

1. <span data-ttu-id="3dfc0-226">[タスク フィールド **の取得] ボタンを** 選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-226">Select the **Get Task Fields** button.</span></span> <span data-ttu-id="3dfc0-227">関数は、タスク名、インデックス、開始日、期間、優先度、およびタスクノートに対して関数を複数回 `getTaskFields` `getTaskfieldAsync` 呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-227">The `getTaskFields` function calls the `getTaskfieldAsync` function multiple times for the task name, index, start date, duration, priority, and task notes.</span></span>

    - <span data-ttu-id="3dfc0-228">名前: `T2`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-228">Name: `T2`</span></span>
    - <span data-ttu-id="3dfc0-229">ID: `2`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-229">ID: `2`</span></span>
    - <span data-ttu-id="3dfc0-230">開始: `Thu 6/14/12`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-230">Start: `Thu 6/14/12`</span></span>
    - <span data-ttu-id="3dfc0-231">期間: `4d`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-231">Duration: `4d`</span></span>
    - <span data-ttu-id="3dfc0-232">優先度: `500`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-232">Priority: `500`</span></span>
    - <span data-ttu-id="3dfc0-233">ノート: これは、タスク T2 のノートです。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-233">Notes: This is a note for task T2.</span></span> <span data-ttu-id="3dfc0-234">単なるテスト ノートです。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-234">It is only a test note.</span></span> <span data-ttu-id="3dfc0-235">実際のノートの場合は、実際の情報になります。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-235">If it had been a real note, there would be some real information.</span></span>

1. <span data-ttu-id="3dfc0-p141">**[getWSSUrlAsync]** ボタンを選択します。プロジェクトが次の種類のどちらかであれば、タスク リストの URL と名前が結果に表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p141">Select the **getWSSUrlAsync** button. If the project is one of the following kinds, the results show the task list URL and name.</span></span>

    - <span data-ttu-id="3dfc0-238">Project Server にインポートされた SharePoint タスク リスト</span><span class="sxs-lookup"><span data-stu-id="3dfc0-238">A SharePoint task list that was imported to Project Server.</span></span>
    - <span data-ttu-id="3dfc0-239">Project Professional にインポートされ、SharePoint に (Project Server を使用せずに) 保存された SharePoint タスク リスト</span><span class="sxs-lookup"><span data-stu-id="3dfc0-239">A SharePoint task list that was imported to Project Professional, and then saved back in SharePoint (not using Project Server).</span></span>

    > [!NOTE]
    > <span data-ttu-id="3dfc0-240">Project Professional が Windows Server コンピューターにインストールされており、プロジェクトを SharePoint に保存できる場合は、**サーバー マネージャー** を使用して **デスクトップ エクスペリエンス** 機能を追加できます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-240">If Project Professional is installed on a Windows Server computer, to be able to save the project back to SharePoint, you can use the **Server Manager** to add the **Desktop Experience** feature.</span></span>

    <span data-ttu-id="3dfc0-241">プロジェクトがローカル プロジェクトの場合、または Project Professional を使用して Project Server によって管理されているプロジェクトを開く場合、メソッドは未定義のエラー `getWSSUrlAsync` を表示します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-241">If the project is a local project, or if you use Project Professional to open a project that is managed by Project Server, the `getWSSUrlAsync` method shows an undefined error.</span></span>

    - <span data-ttu-id="3dfc0-242">SharePoint URL: `http://ServerName`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-242">SharePoint URL: `http://ServerName`</span></span>
    - <span data-ttu-id="3dfc0-243">リスト名: `Test task list`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-243">List name: `Test task list`</span></span>

1. <span data-ttu-id="3dfc0-244">**TaskSelectionChanged** イベント セクションの [追加] ボタンを選択します。このセクションでは、関数を呼び出してタスク選択変更イベントを登録し、テキスト ボックス `manageTaskEventHandler` `In onComplete function for addHandlerAsync Status: succeeded` に戻します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-244">Select the **Add** button in the **TaskSelectionChanged event** section, which calls the `manageTaskEventHandler` function to register a task selection changed event and returns `In onComplete function for addHandlerAsync Status: succeeded` in the text box.</span></span> <span data-ttu-id="3dfc0-245">別のタスクを選択します。テキスト ボックスには、 `In task selection changed event handler` タスク選択変更イベントのコールバック関数の出力が表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-245">Select a different task; the text box shows `In task selection changed event handler`, which is the output of the callback function for the task selection changed event.</span></span> <span data-ttu-id="3dfc0-246">イベント ハンドラーの **登録を** 解除するには、[削除] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-246">Choose the **Remove** button to unregister the event handler.</span></span>

1. <span data-ttu-id="3dfc0-247">リソースに関するメソッドを使用するには、最初に [**リソース シート**]、[**リソース配分状況**]、[**リソース フォーム**] などのビューを選択し、次にそのビュー内でリソースを選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-247">To use the resource methods, first select a view such as **Resource Sheet**, **Resource Usage**, or **Resource Form**, and then select a resource in that view.</span></span> <span data-ttu-id="3dfc0-248">**resourceGuid 変数を初期化するには、getSelectedResourceAsync** を選択し、[リソース フィールドの取得] を選択して、リソース プロパティを複数回呼び `getResourceFieldAsync` 出します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-248">Choose **getSelectedResourceAsync** to initialize the **resourceGuid** variable, and then choose **Get Resource Fields** to call `getResourceFieldAsync` multiple times for the resource properties.</span></span> <span data-ttu-id="3dfc0-249">また、リソース選択変更のイベント ハンドラーを追加または削除することもできます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-249">You can also add or remove the resource selection changed event handler.</span></span>

    - <span data-ttu-id="3dfc0-250">リソース名: `R1`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-250">Resource name: `R1`</span></span>
    - <span data-ttu-id="3dfc0-251">原価: `$800.00`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-251">Cost: `$800.00`</span></span>
    - <span data-ttu-id="3dfc0-252">標準単価: `$50.00/h`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-252">Standard Rate: `$50.00/h`</span></span>
    - <span data-ttu-id="3dfc0-253">実績コスト: `$0.00`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-253">Actual Cost: `$0.00`</span></span>
    - <span data-ttu-id="3dfc0-254">実績作業時間 : `0h`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-254">Actual Work: `0h`</span></span>
    - <span data-ttu-id="3dfc0-255">単位: `100%`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-255">Units: `100%`</span></span>

1. <span data-ttu-id="3dfc0-256">アクティブ **なビューの種類と名前を表示するには、[getSelectedViewAsync]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-256">Select **getSelectedViewAsync** to show the type and name of the active view.</span></span> <span data-ttu-id="3dfc0-257">また、ビュー選択変更のイベント ハンドラーを追加または削除することもできます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-257">You can also add or remove the view selection changed event handler.</span></span> <span data-ttu-id="3dfc0-258">たとえば、リソース フォーム **がアクティブ** ビューの場合、関数 `getSelectedViewAsync` はテキスト ボックスに次の情報を表示します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-258">For example, if **Resource Form** is the active view, the `getSelectedViewAsync` function shows the following in the text box.</span></span>

    - <span data-ttu-id="3dfc0-259">ビューの種類: `6`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-259">View type: `6`</span></span>
    - <span data-ttu-id="3dfc0-260">名前: `Resource Form`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-260">Name: `Resource Form`</span></span>

1. <span data-ttu-id="3dfc0-261">[Get **Project フィールド] を選択** して、アクティブなプロジェクトの異なるプロパティに対して関数 `getProjectFieldAsync` を複数回呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-261">Select **Get Project Fields** to call the `getProjectFieldAsync` function multiple times for different properties of the active project.</span></span> <span data-ttu-id="3dfc0-262">プロジェクトが新しいインスタンスから開Project Web App、関数はインスタンス `getProjectFieldAsync` の URL をProject Web Appできます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-262">If the project is opened from Project Web App, the `getProjectFieldAsync` function can get the URL of the Project Web App instance.</span></span>

    - <span data-ttu-id="3dfc0-263">プロジェクト GUID: `9845922E-DAB4-E111-8AF3-00155D3BA208`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-263">Project GUID: `9845922E-DAB4-E111-8AF3-00155D3BA208`</span></span>
    - <span data-ttu-id="3dfc0-264">開始: `Tue 6/12/12`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-264">Start: `Tue 6/12/12`</span></span>
    - <span data-ttu-id="3dfc0-265">終了: `Tue 6/19/12`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-265">Finish: `Tue 6/19/12`</span></span>
    - <span data-ttu-id="3dfc0-266">通貨桁数: `2`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-266">Currency digits: `2`</span></span>
    - <span data-ttu-id="3dfc0-267">通貨記号: `$`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-267">Currency symbol: `$`</span></span>
    - <span data-ttu-id="3dfc0-268">記号の位置: `0`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-268">Symbol position: `0`</span></span>
    - <span data-ttu-id="3dfc0-269">Project Web App の URL: `http://servername/pwa`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-269">Project web app URL: `http://servername/pwa`</span></span>
  
1. <span data-ttu-id="3dfc0-270">[コンテキスト値 **の** 取得] ボタンを選択すると、Office.Context.doc **ument** オブジェクトとオブジェクトのプロパティを取得して、アドインが実行されているドキュメントとアプリケーションのプロパティを取得 `Office.context.application` します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-270">Select the **Get Context Values** button get properties of the document and the application in which the add-in is running, by getting properties of the **Office.Context.document** object and the `Office.context.application` object.</span></span> <span data-ttu-id="3dfc0-271">For example, if the Project1.mpp file is on the local computer desktop, the document URL is `C:\Users\UserAlias\Desktop\Project1.mpp`.</span><span class="sxs-lookup"><span data-stu-id="3dfc0-271">For example, if the Project1.mpp file is on the local computer desktop, the document URL is `C:\Users\UserAlias\Desktop\Project1.mpp`.</span></span> <span data-ttu-id="3dfc0-272">If the .mpp file is in a SharePoint library, the value is the URL of the document.</span><span class="sxs-lookup"><span data-stu-id="3dfc0-272">If the .mpp file is in a SharePoint library, the value is the URL of the document.</span></span> <span data-ttu-id="3dfc0-273">If you use Project Professional 2013 to open a project named Project1 from Project Web App, the document URL is  `<>\Project1`.</span><span class="sxs-lookup"><span data-stu-id="3dfc0-273">If you use Project Professional 2013 to open a project named Project1 from Project Web App, the document URL is  `<>\Project1`.</span></span>

    - <span data-ttu-id="3dfc0-274">ドキュメントの URL: `<>\Project1`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-274">Document URL: `<>\Project1`</span></span>
    - <span data-ttu-id="3dfc0-275">ドキュメント モード: `readWrite`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-275">Document mode: `readWrite`</span></span>
    - <span data-ttu-id="3dfc0-276">アプリの言語: `en-US`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-276">App language: `en-US`</span></span>
    - <span data-ttu-id="3dfc0-277">表示言語: `en-US`</span><span class="sxs-lookup"><span data-stu-id="3dfc0-277">Display language: `en-US`</span></span>

1. <span data-ttu-id="3dfc0-p147">ソース コードを編集した後は、Project をいったん閉じて再起動することで、アドインを最新の情報に更新できます。[**プロジェクト**] リボンの [**Office アドイン**] ドロップダウン リストに、最近使用したアドインの一覧が保持されています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p147">You can refresh the add-in after you edit the source code by closing and restarting Project. In the **Project** ribbon, the **Office Add-ins** drop-down list maintains the list of recently used add-ins.</span></span>

## <a name="example"></a><span data-ttu-id="3dfc0-280">例</span><span class="sxs-lookup"><span data-stu-id="3dfc0-280">Example</span></span>

<span data-ttu-id="3dfc0-p148">Project 2013 SDK のダウンロードには、JSOMCall.html ファイル、JSOM_Sample.js ファイル、関連する Office.js、Office.debug.js、Project-15.js、および Project-15.debug.js の各ファイルの完全なコードが含まれています。次に、JSOMCall.html ファイルのコードを示します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p148">The Project 2013 SDK download contains the complete code in the JSOMCall.html file, the JSOM_Sample.js file, and the related Office.js, Office.debug.js, Project-15.js, and Project-15.debug.js files. Following is the code in the JSOMCall.html file.</span></span>

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

## <a name="robust-programming"></a><span data-ttu-id="3dfc0-283">堅牢なプログラミング</span><span class="sxs-lookup"><span data-stu-id="3dfc0-283">Robust programming</span></span>

<span data-ttu-id="3dfc0-284">**OM Test アドインProject** は、Project 2013 の JavaScript 関数の一部を Project-15.js および Office.js ファイルで使用する例です。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-284">The **Project OM Test** add-in is an example that shows the use of some JavaScript functions for Project 2013 in the Project-15.js and Office.js files.</span></span> <span data-ttu-id="3dfc0-285">この例は単なるテスト用で、堅牢なエラー チェックは含まれていません。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-285">The example is for testing only and does not include robust error checks.</span></span> <span data-ttu-id="3dfc0-286">たとえば、リソースを選択して関数を実行しない場合、変数は初期化され、エラーを返 `getSelectedResourceAsync` `resourceGuid` `getResourceFieldAsync` す呼び出しが行われます。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-286">For example, if you do not select a resource and run the `getSelectedResourceAsync` function, the `resourceGuid` variable is not initialized, and calls to `getResourceFieldAsync` return an error.</span></span> <span data-ttu-id="3dfc0-287">実際に運用するアドインでは、特定のエラーをチェックして結果を無視したり、特定の状況に該当しない機能を隠したり、機能を使用する前にビューや有効な項目を選択するようにユーザーに通知したりする必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-287">For a production add-in, you should check for specific errors and ignore the results, hide functionality that does not apply, or notify the user to choose a view and make a valid selection before using a function.</span></span>

<span data-ttu-id="3dfc0-288">簡単な例では、次のコードのエラー出力には、関数のエラーを回避するために実行するアクションを指定する th 変数  `actionMessage` が含 `getSelectedResourceAsync` まれています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-288">For a simple example, the error output in the following code includes th  `actionMessage` variable that specifies the action to take to avoid an error in the `getSelectedResourceAsync` function.</span></span>

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

<span data-ttu-id="3dfc0-289">Project 2013 SDK のダウンロードの **HelloProject_OData** サンプルには、JQuery ライブラリを使用してポップアップ エラー メッセージを表示する SurfaceErrors.js ファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-289">The **HelloProject_OData** sample in the Project 2013 SDK download includes the SurfaceErrors.js file that uses the JQuery library to display a pop-up error message.</span></span> <span data-ttu-id="3dfc0-290">図 4 に、"toast" 通知のエラー メッセージを示します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-290">Figure 4 shows the error message in a "toast" notification.</span></span>

<span data-ttu-id="3dfc0-291">次のコードは、SurfaceErrors.jsオブジェクトを作成  `throwError` する th 関数を含 `Toast` むファイルです。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-291">The following code in the SurfaceErrors.js file includes th  `throwError` function that creates a `Toast` object.</span></span>

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

<span data-ttu-id="3dfc0-292">この関数を使用するには、JQuery ライブラリと SurfaceErrors.js スクリプトを JSOMCall.html ファイルに含め、他の JavaScript 関数 (など) に呼び出しを追加 `throwError` `throwError` します `logMethodError` 。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-292">To use the `throwError` function, include the JQuery library and the SurfaceErrors.js script in the JSOMCall.html file, and then add a call to `throwError` in other JavaScript functions such as `logMethodError`.</span></span>

> [!NOTE]
> <span data-ttu-id="3dfc0-p151">アドインを展開する前に、office.js の参照と jQuery の参照をコンテンツ配信ネットワーク (CDN) の参照に変更してください。CDN の参照は最新のバージョンと高いパフォーマンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3dfc0-p151">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

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

<br/>

```js
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```

<br/>

<span data-ttu-id="3dfc0-295">*図 4. SurfaceErrors.js ファイル内の関数は "toast" 通知を表示できます*</span><span class="sxs-lookup"><span data-stu-id="3dfc0-295">*Figure 4. Functions in the SurfaceErrors.js file can show a "toast" notification*</span></span>

![SurfaceError ルーチンを使用してエラーを表示する。](../images/pj15-create-simple-agave-surface-error.png)


## <a name="see-also"></a><span data-ttu-id="3dfc0-297">関連項目</span><span class="sxs-lookup"><span data-stu-id="3dfc0-297">See also</span></span>

- [<span data-ttu-id="3dfc0-298">Project 用の作業ウィンドウ アドイン</span><span class="sxs-lookup"><span data-stu-id="3dfc0-298">Task pane add-ins for Project</span></span>](../project/project-add-ins.md)
- [<span data-ttu-id="3dfc0-299">アドイン用の JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="3dfc0-299">Understanding the JavaScript API for add-ins</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="3dfc0-300">OfficeJavaScript API アドイン</span><span class="sxs-lookup"><span data-stu-id="3dfc0-300">Office JavaScript API Add-ins</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="3dfc0-301">Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="3dfc0-301">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="3dfc0-302">Project 2013 SDK のダウンロード</span><span class="sxs-lookup"><span data-stu-id="3dfc0-302">Project 2013 SDK download</span></span>](https://www.microsoft.com/download/details.aspx?id=30435%20)
