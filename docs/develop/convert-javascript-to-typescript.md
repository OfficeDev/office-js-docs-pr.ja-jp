---
title: Visual Studio の Office アドイン プロジェクトを TypeScript に変換する
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 894cdcb8360a26dfb0f2d5ddbf06cbd7c52d5623
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945349"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="ba4a7-102">Visual Studio の Office アドイン プロジェクトを TypeScript に変換する</span><span class="sxs-lookup"><span data-stu-id="ba4a7-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="ba4a7-103">Visual Studio の Office アドイン テンプレートを使用して JavaScript を使用するアドインを作成すると、そのアドイン プロジェクトは TypeScript に変換できます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="ba4a7-104">アドイン プロジェクトの作成に Visual Studio を使用することで、Office アドイン TypeScript プロジェクトをゼロから作成する必要がなくなります。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-104">By using Visual Studio to create the add-in project, you avoid having to create your Office Add-in TypeScript project from scratch.</span></span> 

<span data-ttu-id="ba4a7-105">この記事では、Visual Studio を使用して Excel アドインを作成して、そのアドイン プロジェクトを JavaScript から TypeScript に変換する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-105">This article shows you how to create an Excel add-in using Visual Studio and then convert the add-in project from JavaScript to TypeScript.</span></span> <span data-ttu-id="ba4a7-106">同じ手順を使用すると、その他の種類の Office アドイン JavaScript プロジェクトを Visual Studio の TypeScript に変換できます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-106">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="ba4a7-107">Visual Studio を使用することなく Office アドイン TypeScript プロジェクトを作成するには、「[5 分間クイック スタート](../index.yml)」の「任意のエディタ」セクションに示された手順を実行して、[Office アドイン用 Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)のプロンプトが表示されたら `TypeScript` を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ba4a7-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="ba4a7-108">Prerequisites</span></span>

- <span data-ttu-id="ba4a7-109">**Office/SharePoint 開発**ワークロードがインストールされている [Visual Studio 2017](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="ba4a7-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="ba4a7-110">既に Visual Studio 2017 がインストールされている場合は、[Visual Studio インストーラー](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)を使用して、**Office/SharePoint 開発**ワークロードがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> 

- <span data-ttu-id="ba4a7-111">Visual Studio 2017 用 TypeScript 2.3</span><span class="sxs-lookup"><span data-stu-id="ba4a7-111">TypeScript 2.3 for Visual Studio 2017</span></span>

    > [!NOTE]
    > <span data-ttu-id="ba4a7-112">TypeScript は、既定で Visual Studio 2017 と共にインストールされますが、TypeScript がインストールされているかどうかは、[Visual Studio インストーラーを使用して](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)確認できます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-112">TypeScript should be installed by default with Visual Studio 2017, but you can [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to confirm that it is installed.</span></span> <span data-ttu-id="ba4a7-113">Visual Studio インストーラーで、**[個別のコンポーネント]** タブを選択して、**[SDK、ライブラリ、およびフレームワーク]** の下で **[TypeScript 2.3 SDK]** が選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-113">In the Visual Studio Installer, select the **Individual components** tab and then verify that **TypeScript 2.3 SDK** is selected under **SDKs, libraries, and frameworks**.</span></span>

- <span data-ttu-id="ba4a7-114">Excel 2016 以降</span><span class="sxs-lookup"><span data-stu-id="ba4a7-114">Excel 2016, version 6769.2011 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="ba4a7-115">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="ba4a7-115">Create the add-in project</span></span>

1. <span data-ttu-id="ba4a7-116">Visual Studio を開いて、Visual Studio のメニュー バーから、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-116">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="ba4a7-117">**[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[Excel Web アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-117">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="ba4a7-118">プロジェクトに名前を付けて、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-118">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="ba4a7-119">**[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を Excel に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-119">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="ba4a7-p104">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="ba4a7-122">アドイン プロジェクトを TypeScript に変換する</span><span class="sxs-lookup"><span data-stu-id="ba4a7-122">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="ba4a7-123">**ソリューション エクスプローラー**で、**Home.js** ファイルの名前を **Home.ts** に変更します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-123">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ba4a7-p105">TypeScript プロジェクトには、TypeScript ファイルと JavaScript ファイルをどちらも一緒に含めることができ、プロジェクトはコンパイルされます。TypeScript は、JavaScript にコンパイルされる JavaScript の型付けスーパーセットであるためです。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="ba4a7-126">ファイル名拡張子の変更を確認するダイアログが表示されたら、**[はい]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-126">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="ba4a7-127">Web アプリケーション プロジェクトのルートに **Office.d.ts** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-127">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="ba4a7-128">Web ブラウザーで、[Office.js の型定義ファイル](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)を開きます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-128">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="ba4a7-129">このファイルの内容をクリップボードにコピーします。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-129">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="ba4a7-130">Visual Studio で、**Office.d.ts** ファイルを開きます。このファイルにクリップボードの内容を貼り付けてから、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-130">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="ba4a7-131">Web アプリケーション プロジェクトのルートに、**jQuery.d.ts** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-131">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="ba4a7-132">Web ブラウザーで、[jQuery の型定義ファイル](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts)を開きます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-132">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span></span> <span data-ttu-id="ba4a7-133">このファイルの内容をクリップボードにコピーします。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-133">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="ba4a7-134">Visual Studio で、**jQuery.d.ts** ファイルを開きます。このファイルにクリップボードの内容を貼り付けてから、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-134">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="ba4a7-135">Visual Studio で、Web アプリケーション プロジェクトのルートに **tsconfig.json** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-135">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="ba4a7-136">**tsconfig.json** ファイルを開いて、次の内容をファイルに追加してから、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-136">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="ba4a7-137">**Home.ts** ファイルを開いて、次の宣言をファイルの先頭に追加します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-137">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```javascript
    declare var fabric: any;
    ```

12. <span data-ttu-id="ba4a7-138">**Home.ts** ファイルで、次に示す行の **'1.1'** を **1.1** に変更して (つまり、引用符を削除して)、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-138">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="ba4a7-139">変換後のアドイン プロジェクトを実行する</span><span class="sxs-lookup"><span data-stu-id="ba4a7-139">Run the converted add-in project</span></span>

1. <span data-ttu-id="ba4a7-p108">Visual Studio で、F5 キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Excel を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="ba4a7-142">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-142">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="ba4a7-143">ワークシートで、数値を格納している 9 つのセルを選択します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-143">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="ba4a7-144">作業ウィンドウの **[強調表示]** ボタンをクリックして、選択した範囲内で最大の値を格納しているセルを強調表示にします。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-144">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="ba4a7-145">Home.ts コード ファイル</span><span class="sxs-lookup"><span data-stu-id="ba4a7-145">Home.ts code file</span></span>

<span data-ttu-id="ba4a7-146">参考のために、次のコード スニペットで、これまでに説明した変更点を適用した後の **Home.ts** ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-146">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="ba4a7-147">このコードには、アドインを実行するために必要な最小限の変更点が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ba4a7-147">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```javascript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016 or later, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
```

## <a name="see-also"></a><span data-ttu-id="ba4a7-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="ba4a7-148">See also</span></span>

* [<span data-ttu-id="ba4a7-149">StackOverflow における Promise 実装に関するディスカッション</span><span class="sxs-lookup"><span data-stu-id="ba4a7-149">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="ba4a7-150">GitHub における Office アドインのサンプル</span><span class="sxs-lookup"><span data-stu-id="ba4a7-150">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
