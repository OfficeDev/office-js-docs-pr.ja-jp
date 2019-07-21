---
title: Visual Studio の Office アドイン プロジェクトを TypeScript に変換する
description: ''
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: 7c51479c1a5d1df5d9b0622dbae4fe9f01ad0c2c
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771307"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="335a3-102">Visual Studio の Office アドイン プロジェクトを TypeScript に変換する</span><span class="sxs-lookup"><span data-stu-id="335a3-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="335a3-103">Visual Studio の Office アドイン テンプレートを使用して JavaScript を使用するアドインを作成すると、そのアドイン プロジェクトは TypeScript に変換できます。</span><span class="sxs-lookup"><span data-stu-id="335a3-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="335a3-104">この記事では、Excel アドイン用のこの変換プロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="335a3-104">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="335a3-105">同じ手順を使用すると、その他の種類の Office アドイン プロジェクトを JavaScript から Visual Studio の TypeScript に変換できます。</span><span class="sxs-lookup"><span data-stu-id="335a3-105">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="335a3-106">Visual Studio を使用することなく Office アドイン TypeScript プロジェクトを作成するには、「[5 分間のクイック スタート](../index.md)」の「Yeoman ジェネレーター」のセクションに示された手順を実行して、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)のプロンプトが表示されたら `TypeScript` を選択します。</span><span class="sxs-lookup"><span data-stu-id="335a3-106">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quick start](../index.md) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="335a3-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="335a3-107">Prerequisites</span></span>

- <span data-ttu-id="335a3-108">**Office/SharePoint 開発**ワークロードがインストールされている [Visual Studio 2017](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="335a3-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="335a3-109">既に Visual Studio 2017 がインストールされている場合は、[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)を使用して、**Office/SharePoint 開発**ワークロードがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="335a3-109">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="335a3-110">このワークロードがまだインストールされていない場合は、Visual Studio インストーラーを使用して[インストール](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads)してください。</span><span class="sxs-lookup"><span data-stu-id="335a3-110">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span></span>

- <span data-ttu-id="335a3-111">TypeScript SDK バージョン 2.3 以降 (Visual Studio 2017 用)</span><span class="sxs-lookup"><span data-stu-id="335a3-111">TypeScript SDK version 2.3 or later (for Visual Studio 2017)</span></span>

    > [!TIP]
    > <span data-ttu-id="335a3-112">[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)で、**[個別のコンポーネント]** タブを選択して、**[SDK、ライブラリ、およびフレームワーク]** セクションまでスクロール ダウンします。</span><span class="sxs-lookup"><span data-stu-id="335a3-112">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="335a3-113">そのセクション内で、**TypeScript SDK** コンポーネント (バージョン 2.3 以降) のうち少なくとも 1 つが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="335a3-113">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="335a3-114">**TypeScript SDK** コンポーネントが選択されていない場合は、使用可能な最新バージョンの SDK を選択し、**[変更]** ボタンを選択して、[個々のコンポーネントをインストール](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components)します。</span><span class="sxs-lookup"><span data-stu-id="335a3-114">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components).</span></span> 

- <span data-ttu-id="335a3-115">Excel 2016 以降</span><span class="sxs-lookup"><span data-stu-id="335a3-115">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="335a3-116">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="335a3-116">Create the add-in project</span></span>

1. <span data-ttu-id="335a3-117">Visual Studio を開いて、Visual Studio のメニュー バーから、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="335a3-117">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="335a3-118">**[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[Excel Web アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="335a3-118">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="335a3-119">プロジェクトに名前を付けて、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="335a3-119">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="335a3-120">**[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を Excel に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="335a3-120">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="335a3-p104">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="335a3-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="335a3-123">アドイン プロジェクトを TypeScript に変換する</span><span class="sxs-lookup"><span data-stu-id="335a3-123">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="335a3-124">**ソリューション エクスプローラー**で、**Home.js** ファイルの名前を **Home.ts** に変更します。</span><span class="sxs-lookup"><span data-stu-id="335a3-124">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="335a3-p105">TypeScript プロジェクトには、TypeScript ファイルと JavaScript ファイルをどちらも一緒に含めることができ、プロジェクトはコンパイルされます。TypeScript は、JavaScript にコンパイルされる JavaScript の型付けスーパーセットであるためです。</span><span class="sxs-lookup"><span data-stu-id="335a3-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="335a3-127">ファイル名拡張子の変更を確認するダイアログが表示されたら、**[はい]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="335a3-127">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="335a3-128">Web アプリケーション プロジェクトのルートに **Office.d.ts** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="335a3-128">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="335a3-p106">Web ブラウザーで、[Office.js の型定義ファイル](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)を開きます。 このファイルの内容をクリップボードにコピーします。</span><span class="sxs-lookup"><span data-stu-id="335a3-p106">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="335a3-131">Visual Studio で、**Office.d.ts** ファイルを開きます。このファイルにクリップボードの内容を貼り付けてから、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="335a3-131">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="335a3-132">Web アプリケーション プロジェクトのルートに、**jQuery.d.ts** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="335a3-132">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="335a3-p107">Web ブラウザーで、[jQuery の型定義ファイル](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts)を開きます。 このファイルの内容をクリップボードにコピーします。</span><span class="sxs-lookup"><span data-stu-id="335a3-p107">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts). Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="335a3-135">Visual Studio で、**jQuery.d.ts** ファイルを開きます。このファイルにクリップボードの内容を貼り付けてから、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="335a3-135">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="335a3-136">Visual Studio で、Web アプリケーション プロジェクトのルートに **tsconfig.json** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="335a3-136">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="335a3-137">**tsconfig.json** ファイルを開いて、次の内容をファイルに追加してから、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="335a3-137">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```json
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ],
            "sourceMap": true
        }
    }
    ```

11. <span data-ttu-id="335a3-138">**Home.ts** ファイルを開いて、次の宣言をファイルの先頭に追加します。</span><span class="sxs-lookup"><span data-stu-id="335a3-138">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```typescript
    declare var fabric: any;
    ```

12. <span data-ttu-id="335a3-139">**Home.ts** ファイルで、行 `Office.initialize = function (reason) {` を見つけます。その直後に一行追加して、ここに示されているようにグローバル `window.Promise` をポリフィルします。</span><span class="sxs-lookup"><span data-stu-id="335a3-139">In the **Home.ts** file, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```typescript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

13. <span data-ttu-id="335a3-140">**Home.ts** ファイルで、次に示す行の **'1.1'** を **1.1** に変更します (つまり、引用符を削除します)。</span><span class="sxs-lookup"><span data-stu-id="335a3-140">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line:</span></span>

    ```typescript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

14. <span data-ttu-id="335a3-141">**Home.ts** ファイルで、`displaySelectedCells` 関数を検索し、関数全体を次のコードで置換し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="335a3-141">In the **Home.ts** file, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

    ```typescript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="335a3-142">変換後のアドイン プロジェクトを実行する</span><span class="sxs-lookup"><span data-stu-id="335a3-142">Run the converted add-in project</span></span>

1. <span data-ttu-id="335a3-143">Visual Studio で、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Excel を起動します。</span><span class="sxs-lookup"><span data-stu-id="335a3-143">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="335a3-144">アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="335a3-144">The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="335a3-145">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="335a3-145">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="335a3-146">ワークシートで、数値を格納している 9 つのセルを選択します。</span><span class="sxs-lookup"><span data-stu-id="335a3-146">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="335a3-147">作業ウィンドウの **[強調表示]** ボタンをクリックして、選択した範囲内で最大の値を格納しているセルを強調表示にします。</span><span class="sxs-lookup"><span data-stu-id="335a3-147">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="335a3-148">Home.ts コード ファイル</span><span class="sxs-lookup"><span data-stu-id="335a3-148">Home.ts code file</span></span>

<span data-ttu-id="335a3-p109">参考のために、次のコード スニペットで、これまでに説明した変更点を適用した後の **Home.ts** ファイルの内容を示します。 このコードには、アドインを実行するために必要な最小限の変更点が含まれています。</span><span class="sxs-lookup"><span data-stu-id="335a3-p109">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```typescript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;

        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
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
            null,
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

## <a name="see-also"></a><span data-ttu-id="335a3-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="335a3-151">See also</span></span>

* [<span data-ttu-id="335a3-152">StackOverflow における Promise 実装に関するディスカッション</span><span class="sxs-lookup"><span data-stu-id="335a3-152">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="335a3-153">GitHub における Office アドインのサンプル</span><span class="sxs-lookup"><span data-stu-id="335a3-153">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
