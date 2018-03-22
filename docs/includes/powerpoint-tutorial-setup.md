開発プロジェクトを設定して、このチュートリアルを始めます。 

> [!NOTE]
> このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="prerequisites"></a>前提条件

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a>セットアップ

このチュートリアルでは、Visual Studio を使用してアドインを作成します。

### <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

1. [Visual Studio] メニュー バーで、**[ファイル]** > **[新規作成]** > **[プロジェクト]** の順に選択します。
    
2. **[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[PowerPoint Web アドイン]** を選択します。 

3. プロジェクトに **HelloWorld** という名前を付けて、**[OK]** ボタンをクリックします。

4. **[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を PowerPoint に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。

5. Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。

     ![PowerPoint チュートリアル - HelloWorld ソリューションで 2 つのプロジェクトを表示する Visual Studio ソリューション エクスプローラー ウィンドウ](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a>Visual Studio ソリューションについて理解する

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a>コードを更新する 

アドイン コードを次のように編集し、このチュートリアルの後続の手順でアドイン機能を実装するために使用するフレームワークを作成します。

1. **Home.html** では、アドインの作業ウィンドウにレンダリングされる HTML を指定します。 **Home.html** において、`id="content-main"` で **div** を検索し、**div** 全体を次のマークアップと置き換えてファイルを保存します。

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. Web アプリケーション プロジェクトのルートにあるファイル **Home.js** を開きます。 このファイルは、アドイン用のスクリプトを指定します。 すべての内容を次のコードに置き換え、ファイルを保存します。

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```
