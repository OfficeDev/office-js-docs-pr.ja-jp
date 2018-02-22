---
title: Visual Studio の Office アドイン プロジェクトを TypeScript に変換する
description: ''
ms.date: 01/19/2018
---

# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Visual Studio の Office アドイン プロジェクトを TypeScript に変換する

Visual Studio の Office アドイン テンプレートを使用して JavaScript を使用するアドインを作成すると、そのアドイン プロジェクトは TypeScript に変換できます。 アドイン プロジェクトの作成に Visual Studio を使用することで、Office アドイン TypeScript プロジェクトをゼロから作成する必要がなくなります。 

この記事では、Visual Studio を使用して Excel アドインを作成して、そのアドイン プロジェクトを JavaScript から TypeScript に変換する方法について説明します。 同じ手順を使用すると、その他の種類の Office アドイン JavaScript プロジェクトを Visual Studio の TypeScript に変換できます。

> [!NOTE]
> Visual Studio を使用することなく Office アドイン TypeScript プロジェクトを作成するには、「[5 分間のクイック スタート](../index.yml)」の「任意のエディター」のセクションに示された手順を実行して、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)のプロンプトが表示されたら `TypeScript` を選択します。

## <a name="prerequisites"></a>前提条件

- **Office/SharePoint 開発**ワークロードがインストールされている [Visual Studio 2017](https://www.visualstudio.com/vs/)

    > [!NOTE]
    > 既に Visual Studio 2017 がインストールされている場合は、[Visual Studio インストーラー](https://docs.microsoft.com/ja-jp/visualstudio/install/modify-visual-studio)を使用して、**Office/SharePoint 開発**ワークロードがインストールされていることを確認してください。 

- TypeScript 2.3 for Visual Studio 2017

    > [!NOTE]
    > TypeScript は、既定で Visual Studio 2017 と共にインストールされますが、TypeScript がインストールされているかどうかは、[Visual Studio インストーラーを使用して](https://docs.microsoft.com/ja-jp/visualstudio/install/modify-visual-studio)確認できます。 Visual Studio インストーラーで、**[個別のコンポーネント]** タブを選択して、**[SDK、ライブラリ、およびフレームワーク]** の下で **[TypeScript 2.3 SDK]** が選択されていることを確認します。

- Excel 2016

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

1. Visual Studio を開いて、Visual Studio のメニュー バーから、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。

2. **[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[Excel Web アドイン]** を選択します。 

3. プロジェクトに名前を付けて、**[OK]** を選択します。

4. **[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を Excel に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。

5. ソリューションが Visual Studio によって作成され、2 つのプロジェクトが**ソリューション エクスプローラー**に表示されます。 Visual Studio で **Home.html** ファイルが開かれます。

## <a name="convert-the-add-in-project-to-typescript"></a>アドイン プロジェクトを TypeScript に変換する

1. **ソリューション エクスプローラー**で、**Home.js** ファイルの名前を **Home.ts** に変更します。

    > [!NOTE]
    > TypeScript プロジェクトには、TypeScript ファイルと JavaScript ファイルをどちらも一緒に含めることができ、プロジェクトはコンパイルされます。TypeScript は、JavaScript にコンパイルされる JavaScript の型付けスーパーセットであるためです。 

2. ファイル名拡張子の変更を確認するダイアログが表示されたら、**[はい]** を選択します。

3. Web アプリケーション プロジェクトのルートに **Office.d.ts** という名前の新しいファイルを作成します。

4. Web ブラウザーで、[Office.js の型定義ファイル](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)を開きます。 このファイルの内容をクリップボードにコピーします。

5. Visual Studio で、**Office.d.ts** ファイルを開きます。このファイルにクリップボードの内容を貼り付けてから、ファイルを保存します。

6. Web アプリケーション プロジェクトのルートに、**jQuery.d.ts** という名前の新しいファイルを作成します。

7. Web ブラウザーで、[jQuery の型定義ファイル](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts)を開きます。 このファイルの内容をクリップボードにコピーします。

8. Visual Studio で、**jQuery.d.ts** ファイルを開きます。このファイルにクリップボードの内容を貼り付けてから、ファイルを保存します。

9. Visual Studio で、Web アプリケーション プロジェクトのルートに **tsconfig.json** という名前の新しいファイルを作成します。

10. **tsconfig.json** ファイルを開いて、次の内容をファイルに追加してから、ファイルを保存します。

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. **Home.ts** ファイルを開いて、次の宣言をファイルの先頭に追加します。

    ```javascript
    declare var fabric: any;
    ```

12. **Home.ts** ファイルで、次に示す行の **'1.1'** を **1.1** に変更して (つまり、引用符を削除して)、ファイルを保存します。

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a>変換後のアドイン プロジェクトを実行する

1. Visual Studio で、F5 キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Excel を起動します。 アドインは IIS 上でローカルにホストされます。

2. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

3. ワークシートで、数値を格納している 9 つのセルを選択します。

4. 作業ウィンドウの **[強調表示]** ボタンをクリックして、選択した範囲内で最大の値を格納しているセルを強調表示にします。

## <a name="homets-code-file"></a>Home.ts コード ファイル

参考のために、次のコード スニペットで、これまでに説明した変更点を適用した後の **Home.ts** ファイルの内容を示します。 このコードには、アドインを実行するために必要な最小限の変更点が含まれています。

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
            
            // If not using Excel 2016, use fallback logic.
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

## <a name="see-also"></a>関連項目

* [StackOverflow における Promise 実装に関するディスカッション](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [GitHub における Office アドインのサンプル](https://github.com/officedev)
