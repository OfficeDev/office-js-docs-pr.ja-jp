---
title: Visual Studio の Office アドイン プロジェクトを TypeScript に変換する
description: Visual Studio で Office アドイン プロジェクトを変換して TypeScript を使用する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: daa81c3785484083aa49516b04491acad1404884
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889353"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Visual Studio の Office アドイン プロジェクトを TypeScript に変換する

Visual Studio の Office アドイン テンプレートを使用して JavaScript を使用するアドインを作成すると、そのアドイン プロジェクトは TypeScript に変換できます。 この記事では、Excel アドイン用のこの変換プロセスについて説明します。 同じ手順を使用すると、その他の種類の Office アドイン プロジェクトを JavaScript から Visual Studio の TypeScript に変換できます。

> [!IMPORTANT]
> この記事では、F5 キーを押すと、コードが JavaScript に挿入され、自動的に Office にサイドロードされるようにするために必要な *最小限* の手順について説明します。 ただし、コードは "TypeScripty" ではありません。 たとえば、変数はキーワードの代わりに`let`宣言`var`されるか、指定`const`した型で宣言されません。 TypeScript の強力な入力を最大限に活用するには、コードをさらに変更することを検討してください。

> [!NOTE]
> Visual Studio を使用することなく Office アドイン TypeScript プロジェクトを作成するには、「[5 分間のクイック スタート](../index.yml)」の「Yeoman ジェネレーター」のセクションに示された手順を実行して、[Office アドイン用の Yeoman ジェネレーター](yeoman-generator-overview.md)のプロンプトが表示されたら `TypeScript` を選択します。

## <a name="prerequisites"></a>前提条件

- **Office/SharePoint 開発** ワークロードがインストールされた [Visual Studio 2019 以降](https://www.visualstudio.com/vs/)

    > [!TIP]
    > 既に Visual Studio がインストールされている場合は、[Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)を使用して、**Office/SharePoint 開発** ワークロードがインストールされていることを確認してください。 このワークロードがまだインストールされていない場合は、Visual Studio インストーラーを使用して[インストール](/visualstudio/install/modify-visual-studio#modify-workloads)してください。

- TypeScript SDK バージョン 2.3 以降。

    > [!TIP]
    > [Visual Studio インストーラー](/visualstudio/install/modify-visual-studio)で、**[個別のコンポーネント]** タブを選択して、**[SDK、ライブラリ、およびフレームワーク]** セクションまでスクロール ダウンします。 そのセクション内で、**TypeScript SDK** コンポーネント (バージョン 2.3 以降) のうち少なくとも 1 つが選択されていることを確認します。 **TypeScript SDK** コンポーネントが選択されていない場合は、使用可能な最新バージョンの SDK を選択し、[**変更**] を選択して [個々のコンポーネントをインストール](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components)します。

- Excel 2016以降。

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

1. Visual Studio で、[**新しいプロジェクトの作成**] を選択します。

1. 検索ボックスを使用して、**アドイン** と入力します。 [**Excel Web アドイン**] を選択し、[**次へ**] を選択します。

1. プロジェクトに名前を付けて、[**作成**] を選択します。

1. **[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を Excel に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。

1. Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー** に表示されます。**Home.html** ファイルが Visual Studio で開かれます。

## <a name="convert-the-add-in-project-to-typescript"></a>アドイン プロジェクトを TypeScript に変換する

1. **Home.js** ファイルを見つけて、名前を **Home.ts** に変更します。

1. **./Functions/FunctionFile.js** ファイルを見つけて、名前を **FunctionFile.ts** に変更します。

1. **./Scripts/MessageBanner.js** ファイルを見つけて、名前を **MessageBanner.ts** に変更します。

1. [**ツール**] タブから [**NuGet パッケージ マネージャー**] を選択し、[**ソリューション用の NuGet パッケージの管理...**] を選択します。

1. **[参照**] タブを選択した状態で、「**jquery」と入力します。TypeScript.DefinitelyTyped**. このパッケージをインストールするか、既にインストールされている場合は更新します。 これにより、jQuery TypeScript 定義がプロジェクトに確実に含まれます。 jQuery のパッケージは、Visual Studio によって生成されたファイル **(packages.config** と呼ばれます) に表示されます。

    > [!NOTE]
    > TypeScript プロジェクトには、TypeScript ファイルと JavaScript ファイルをどちらも一緒に含めることができ、プロジェクトはコンパイルされます。TypeScript は、JavaScript にコンパイルされる JavaScript の型付けスーパーセットであるためです。

1. **Home.ts** で、次に示すように、行`Office.initialize = function (reason) {`を見つけて、その直後に線を追加してグローバル`window.Promise`をポリフィルします。

    ```TypeScript
    Office.initialize = function (reason) {
        // Add the following line.
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

1. **Home.ts** で関数を`displaySelectedCells`見つけ、関数全体を次のコードに置き換えて、ファイルを保存します。

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

1. **./Scripts/MessageBanner.ts** で、行 `_onResize(null);` を見つけて、次のものに置き換えます。

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a>変換後のアドイン プロジェクトを実行する

1. Visual Studio で、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Excel を起動します。 アドインは IIS 上でローカルにホストされます。

1. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

1. ワークシートで、数値を格納している 9 つのセルを選択します。

1. 作業ウィンドウの **[強調表示]** ボタンをクリックして、選択した範囲内で最大の値を格納しているセルを強調表示にします。

## <a name="homets-code-file"></a>Home.ts コード ファイル

参考のために、次のコード スニペットで、これまでに説明した変更点を適用した後の **Home.ts** ファイルの内容を示します。 このコードには、アドインを実行するために必要な最小限の変更点が含まれています。

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

## <a name="see-also"></a>関連項目

- [StackOverflow における Promise 実装に関するディスカッション](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [GitHub における Office アドインのサンプル](https://github.com/OfficeDev/Office-Add-in-samples)
