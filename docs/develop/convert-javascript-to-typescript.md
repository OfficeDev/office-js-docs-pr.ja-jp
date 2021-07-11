---
title: Visual Studio の Office アドイン プロジェクトを TypeScript に変換する
description: TypeScript を使用するために、Officeアドイン プロジェクトをVisual Studioする方法について説明します。
ms.date: 09/01/2020
localization_priority: Normal
ms.openlocfilehash: 2932c94052735436ce1357966832528679e6900d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349981"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="a4004-103">Visual Studio の Office アドイン プロジェクトを TypeScript に変換する</span><span class="sxs-lookup"><span data-stu-id="a4004-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="a4004-104">Visual Studio の Office アドイン テンプレートを使用して JavaScript を使用するアドインを作成すると、そのアドイン プロジェクトは TypeScript に変換できます。</span><span class="sxs-lookup"><span data-stu-id="a4004-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="a4004-105">この記事では、Excel アドイン用のこの変換プロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="a4004-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="a4004-106">同じ手順を使用すると、その他の種類の Office アドイン プロジェクトを JavaScript から Visual Studio の TypeScript に変換できます。</span><span class="sxs-lookup"><span data-stu-id="a4004-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a4004-107">この記事では、F5 キーを押すとコードが JavaScript にトランスピロードされ、コードが自動的に Office にサイドロードされるのを確認するために必要な最小限の手順について説明します。 </span><span class="sxs-lookup"><span data-stu-id="a4004-107">This article describes the *minimal* steps necessary to ensure that, when you press F5, the code will be transpiled to JavaScript which is then sideloaded automatically into Office.</span></span> <span data-ttu-id="a4004-108">ただし、コードは非常に "TypeScripty" ではありません。</span><span class="sxs-lookup"><span data-stu-id="a4004-108">However, the code is not very "TypeScripty".</span></span> <span data-ttu-id="a4004-109">たとえば、変数はキーワードではなくキーワードで宣言され、指定された `var` `let` 型では宣言されません。</span><span class="sxs-lookup"><span data-stu-id="a4004-109">For example, variables are declared with the `var` keyword instead of `let` and they are not declared with a specified type.</span></span> <span data-ttu-id="a4004-110">TypeScript の強力なタイピングを活用するには、コードをさらに変更する方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="a4004-110">To take full advantage of the strong typing of TypeScript, consider making further changes to the code.</span></span> 

> [!NOTE]
> <span data-ttu-id="a4004-111">Visual Studio を使用することなく Office アドイン TypeScript プロジェクトを作成するには、「[5 分間のクイック スタート](../index.yml)」の「Yeoman ジェネレーター」のセクションに示された手順を実行して、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)のプロンプトが表示されたら `TypeScript` を選択します。</span><span class="sxs-lookup"><span data-stu-id="a4004-111">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a4004-112">前提条件</span><span class="sxs-lookup"><span data-stu-id="a4004-112">Prerequisites</span></span>

- <span data-ttu-id="a4004-113">**Office/SharePoint 開発** ワークロードがインストールされている [Visual Studio 2019](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="a4004-113">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="a4004-114">既に Visual Studio 2019 がインストールされている場合は、[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)を使用して、**Office/SharePoint 開発** ワークロードがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="a4004-114">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="a4004-115">このワークロードがまだインストールされていない場合は、Visual Studio インストーラーを使用して[インストール](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-workloads)してください。</span><span class="sxs-lookup"><span data-stu-id="a4004-115">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-workloads).</span></span>

- <span data-ttu-id="a4004-116">TypeScript SDK バージョン 2.3 以降 (Visual Studio 2019 用)</span><span class="sxs-lookup"><span data-stu-id="a4004-116">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="a4004-117">[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)で、**[個別のコンポーネント]** タブを選択して、**[SDK、ライブラリ、およびフレームワーク]** セクションまでスクロール ダウンします。</span><span class="sxs-lookup"><span data-stu-id="a4004-117">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="a4004-118">そのセクション内で、**TypeScript SDK** コンポーネント (バージョン 2.3 以降) のうち少なくとも 1 つが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a4004-118">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="a4004-119">**TypeScript SDK** コンポーネントが選択されていない場合は、使用可能な最新バージョンの SDK を選択し、**[変更]** ボタンを選択して、[個々のコンポーネントをインストール](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components)します。</span><span class="sxs-lookup"><span data-stu-id="a4004-119">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components).</span></span> 

- <span data-ttu-id="a4004-120">Excel 2016 以降</span><span class="sxs-lookup"><span data-stu-id="a4004-120">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="a4004-121">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="a4004-121">Create the add-in project</span></span>

1. <span data-ttu-id="a4004-122">Visual Studio で、[**新しいプロジェクトの作成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a4004-122">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="a4004-123">検索ボックスを使用して、**アドイン** と入力します。</span><span class="sxs-lookup"><span data-stu-id="a4004-123">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="a4004-124">[**Excel Web アドイン**] を選択し、[**次へ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a4004-124">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="a4004-125">プロジェクトに名前を付けて、[**作成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a4004-125">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="a4004-126">**[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を Excel に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="a4004-126">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="a4004-p106">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー** に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="a4004-p106">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="a4004-129">アドイン プロジェクトを TypeScript に変換する</span><span class="sxs-lookup"><span data-stu-id="a4004-129">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="a4004-130">**Home.js** ファイルを見つけて、名前を **Home.ts** に変更します。</span><span class="sxs-lookup"><span data-stu-id="a4004-130">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="a4004-131">**./Functions/FunctionFile.js** ファイルを見つけて、名前を **FunctionFile.ts** に変更します。</span><span class="sxs-lookup"><span data-stu-id="a4004-131">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="a4004-132">**./Scripts/MessageBanner.js** ファイルを見つけて、名前を **MessageBanner.ts** に変更します。</span><span class="sxs-lookup"><span data-stu-id="a4004-132">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="a4004-133">[**ツール**] タブから [**NuGet パッケージ マネージャー**] を選択し、[**ソリューション用の NuGet パッケージの管理...**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a4004-133">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="a4004-134">[参照] **タブが** 選択されている場合は **、「jquery」と入力します。TypeScript.DefinitelyTyped**.</span><span class="sxs-lookup"><span data-stu-id="a4004-134">With the **Browse** tab selected, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="a4004-135">このパッケージをインストールするか、既にインストールされている場合は更新します。</span><span class="sxs-lookup"><span data-stu-id="a4004-135">Install this package, or update it if it's already installed.</span></span> <span data-ttu-id="a4004-136">これにより、jQuery TypeScript 定義がプロジェクトに確実に含まれます。</span><span class="sxs-lookup"><span data-stu-id="a4004-136">This will ensure the jQuery TypeScript definitions are included in your project.</span></span> <span data-ttu-id="a4004-137">jQuery のパッケージは、Visual Studio という名前のファイル **にpackages.config。**</span><span class="sxs-lookup"><span data-stu-id="a4004-137">The packages for jQuery appear in a file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a4004-p108">TypeScript プロジェクトには、TypeScript ファイルと JavaScript ファイルをどちらも一緒に含めることができ、プロジェクトはコンパイルされます。TypeScript は、JavaScript にコンパイルされる JavaScript の型付けスーパーセットであるためです。</span><span class="sxs-lookup"><span data-stu-id="a4004-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

6. <span data-ttu-id="a4004-140">**Home.ts** で、行 `Office.initialize = function (reason) {` を見つけて、それに続けて一行追加し、次に示されているようにグローバル `window.Promise` をポリフィルします。</span><span class="sxs-lookup"><span data-stu-id="a4004-140">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // Add the following line.
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

7. <span data-ttu-id="a4004-141">**Home.ts** で関数を検索し、関数全体を次のコードに置き換え `displaySelectedCells` 、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="a4004-141">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file.</span></span>

    ```TypeScript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
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

8. <span data-ttu-id="a4004-142">**./Scripts/MessageBanner.ts** で、行 `_onResize(null);` を見つけて、次のものに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a4004-142">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="a4004-143">変換後のアドイン プロジェクトを実行する</span><span class="sxs-lookup"><span data-stu-id="a4004-143">Run the converted add-in project</span></span>

1. <span data-ttu-id="a4004-144">Visual Studio で、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Excel を起動します。</span><span class="sxs-lookup"><span data-stu-id="a4004-144">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="a4004-145">アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="a4004-145">The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="a4004-146">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="a4004-146">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="a4004-147">ワークシートで、数値を格納している 9 つのセルを選択します。</span><span class="sxs-lookup"><span data-stu-id="a4004-147">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="a4004-148">作業ウィンドウの **[強調表示]** ボタンをクリックして、選択した範囲内で最大の値を格納しているセルを強調表示にします。</span><span class="sxs-lookup"><span data-stu-id="a4004-148">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="a4004-149">Home.ts コード ファイル</span><span class="sxs-lookup"><span data-stu-id="a4004-149">Home.ts code file</span></span>

<span data-ttu-id="a4004-p110">参考のために、次のコード スニペットで、これまでに説明した変更点を適用した後の **Home.ts** ファイルの内容を示します。 このコードには、アドインを実行するために必要な最小限の変更点が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a4004-p110">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```typescript
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If you're using Excel 2013, use fallback logic.
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
            $('#highlight-button').click(highlightHighestValue);
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

    function highlightHighestValue() {
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
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
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

## <a name="see-also"></a><span data-ttu-id="a4004-152">関連項目</span><span class="sxs-lookup"><span data-stu-id="a4004-152">See also</span></span>

- [<span data-ttu-id="a4004-153">StackOverflow における Promise 実装に関するディスカッション</span><span class="sxs-lookup"><span data-stu-id="a4004-153">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="a4004-154">GitHub における Office アドインのサンプル</span><span class="sxs-lookup"><span data-stu-id="a4004-154">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)