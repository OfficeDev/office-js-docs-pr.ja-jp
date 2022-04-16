---
title: Script Lab コードからスタンドアロンの Office アドインを作成する
description: スニペットを Script Lab から Yo Office プロジェクトに移動する方法の詳細
ms.topic: how-to
ms.date: 04/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 038d25610e5ef5cc3e4cdbedb2d2a184294c673e
ms.sourcegitcommit: 5ef2c3ed9eb92b56e36c6de77372d3043ad5b021
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/15/2022
ms.locfileid: "64863299"
---
# <a name="create-a-standalone-office-add-in-from-your-script-lab-code"></a>Script Lab コードからスタンドアロンの Office アドインを作成する

Script Lab でスニペットを作成した場合は、それをスタンドアロンのアドインに変換することをお勧めします。 Script Lab から、[Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md) ("Yo Office" とも呼ばれます) によって生成されたプロジェクトにコードをコピーできます。 その後、最終的に他のユーザーに展開できるアドインとしてコードの開発を続けることができます。

この記事の手順は [Visual Studio Code](https://code.visualstudio.com/) を参照していますが、任意のコード エディターを使用できます。

## <a name="create-a-new-yo-office-project"></a>新しい Yo Office プロジェクトを作成する

スニペット コードの新しい開発場所となるスタンドアロンのアドイン プロジェクトを作成する必要があります。

コマンド `yo office --projectType taskpane --ts true --host <host> --name "basic-sample"` を実行します。ここで、`<host>` は次のいずれかの値です。

- Excel
- outlook
- PowerPoint
- Word

> [!IMPORTANT]
> `--name` の引数値は、スペースがない場合でも二重引用符で囲む必要があります。

前のコマンドは、**basic-sample** という名前の新しいプロジェクト フォルダーを作成します。 指定したホストで実行するように構成されており、TypeScript を使用します。 Script Lab は既定で TypeScript を使用しますが、ほとんどのスニペットは JavaScript です。 必要に応じて Yo Office JavaScript プロジェクトをビルドできますが、コピーするコードが JavaScript であることを確認してください。

## <a name="open-the-snippet-in-script-lab"></a>Script Lab でスニペットを開く

Script Lab の既存のスニペットを使用して、Yo Office で生成されたプロジェクトにスニペットをコピーする方法について説明します。

1. Office (Word、Excel、PowerPoint、または Outlook) を開き、Script Lab を開きます。
1. **[Script Lab]** > **[コード]** を選択します。 Outlook で作業している場合は、メール メッセージを開いて、リボンに Script Lab を表示します。
1. Script Lab の作業ウィンドウで、**[サンプル]** を選択します。 次に、作業している Office ホストに基づいて基本的なサンプルを選択します。
    - Excel または Word の場合は、**[基本 API 呼び出し (TypeScript)]** サンプルを選択します。
    - Outlook の場合は、**[アドイン設定の使用]** サンプルを選択します。
    - PowerPoint の場合は、**[基本 API 呼び出し (Office 2013)]** サンプルを選択します。

## <a name="copy-snippet-code-to-visual-studio-code"></a>スニペット コードを Visual Studio コードにコピーする

これで、VS Code のスニペットから Yo Office プロジェクトにコードをコピーできます。

- VS Code で、**basic-sample** プロジェクトを開きます。

次の手順では、Script Lab のいくつかのタブからコードをコピーします。

:::image type="content" source="../images/script-lab-script-tabs.png" alt-text="Script Lab のタブのスクリーンショット。":::

### <a name="copy-task-pane-code"></a>作業ウィンドウ コードをコピーする

1. VS Code で、**/src/taskpane/taskpane.ts** ファイルを開きます。 JavaScript プロジェクトを使用している場合、ファイル名は **taskpane.js** です。
1. Script Lab で、**[スクリプト]** タブを選択します。
1. **[スクリプト]** タブのすべてのコードをクリップボードにコピーします。 **taskpane.ts** (または JavaScript の場合は **taskpane.js**) の内容全体を、コピーしたコードに置き換えます。

### <a name="copy-task-pane-html"></a>作業ウィンドウ HTML をコピーする

1. VS Code で、**/src/taskpane/taskpane.html** ファイルを開きます。
1. Script Lab で、**[HTML]** タブを選択します。
1. **[HTML]** タブのすべての HTML をクリップボードにコピーします。 `<body>` タグ内のすべての HTML をコピーした HTML に置き換えます。

### <a name="copy-task-pane-css"></a>作業ウィンドウ CSS をコピーする

1. VS Code で、**/src/taskpane/taskpane.css** ファイルを開きます。
1. Script Lab で、**[CSS]** タブを選択します。
1. **[CSS]** タブのすべての CSS をクリップボードにコピーします。 **taskpane.css** の内容全体をコピーした CSS に置き換えます。
1. 前の手順で更新したファイルへのすべての変更を保存します。

## <a name="add-jquery-support"></a>jQuery サポートを追加する

Script Lab は、スニペットで jQuery を使用します。 コードを正常に実行するには、この依存関係を Yo Office プロジェクトに追加する必要があります。

1. **taskpane.html** ファイルを開き、次のスクリプト タグを `<head>` セクションに追加します。

    ```html
     <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-3.3.1.js"></script>
    ```

    > [!NOTE]
    > jQuery の特定のバージョンは異なる場合があります。 **[ライブラリ]** タブを選択すると、Script Lab が使用しているバージョンを確認できます。

1. VS Code でターミナルを開き、次のコマンドを入力します。

    ```command&nbsp;line
    npm install --save-dev jquery@3.1.1
    npm install --save-dev @types/jquery@3.3.1
    ```

追加のライブラリ依存関係を持つスニペットを作成した場合は、必ずそれらを Yo Office プロジェクトに追加してください。 Script Lab の **[ライブラリ]** タブですべてのライブラリの依存関係のリストを見つけます。

## <a name="handle-initialization"></a>ハンドルの初期化

Script Lab は、`Office.onReady` の初期化を自動的に処理します。 独自の `Office.onReady` ハンドラーを提供するには、コードを変更する必要があります。

1. **taskpane.ts** (または JavaScript の場合は **taskpane.js**) ファイルを開きます。
1. Excel または Word の場合は、次を置き換えます。

    ```typescript
    $("#run").click(() => tryCatch(run));
    ```

    置換後:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(() => tryCatch(run));
      });
    });
    ```

1. Outlook の場合は、次を置き換えます。

    ```typescript
    $("#get").click(get);
    $("#set").click(set);
    $("#save").click(save);
    ```

    置換後:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#get").click(get);
        $("#set").click(set);
        $("#save").click(save);
      });
    });
    ```

1. PowerPoint の場合は、次を置き換えます。

    ```typescript
    $("#run").click(run);
    ```

    置換後:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(run);
      });
    });
    ```

1. ファイルを保存します。

## <a name="custom-functions"></a>カスタム関数

スニペットでカスタム関数を使用する場合は、Yo Office カスタム関数テンプレートを使用する必要があります。 カスタム関数をスタンドアロンのアドインに変換するには、次の手順に従います。

1. コマンド`yo office --projectType excel-functions --ts true --name "functions-sample"`を実行します。

    > [!IMPORTANT]
    > `--name` の引数値は、スペースがない場合でも二重引用符で囲む必要があります。

1. Excel を開き、次に Script Lab を開きます。
1. **[Script Lab]** > **[コード]** を選択します。
1. Script Lab 作業ウィンドウで、**[サンプル]** を選択してから、**[基本カスタム関数]** サンプルを選択します。
1. **/src/functions/functions.ts** ファイルを開きます。 JavaScript プロジェクトを使用している場合、ファイル名は **functions.js** です。
1. Script Lab で、**[スクリプト]** タブを選択します。
1. **[スクリプト]** タブのすべてのコードをクリップボードにコピーします。 コピーしたコードと共に、**functions.ts** (または JavaScript の場合は **functions.js**) の上部にコードを貼り付けます。
1. ファイルを保存します。

## <a name="test-the-standalone-add-in"></a>スタンドアロン アドインをテストする

すべての手順が完了したら、スタンドアロン アドインを実行してテストします。 次のコマンドを実行して開始します。

```command&nbsp;line
npm start
```

Office が起動し、リボンからアドインの作業ウィンドウを開くことができます。 おめでとうございます! これで、スタンドアロン プロジェクトとしてアドインのビルドを続行できます。

## <a name="console-logging"></a>コンソール ログ

Script Lab の多くのスニペットは、作業ウィンドウの下部にあるコンソール セクションに出力を書き込みます。 Yo Office プロジェクトにはコンソール セクションがありません。 すべての `console.log*` ステートメントは、既定のデバッグ コンソール (ブラウザー開発ツールなど) に書き込みます。 出力を作業ウィンドウに移動する場合は、コードを更新する必要があります。
