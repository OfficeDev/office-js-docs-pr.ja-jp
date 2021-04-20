---
title: Node.js SSO が有効なアドインをカスタマイズする
description: Yeoman ジェネレーターで作成した SSO 対応アドインのカスタマイズについて学習します。
ms.date: 02/01/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 174df5e58e794b94b02025bd90a65f5ae8e26d44
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234171"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a>Node.js SSO が有効なアドインをカスタマイズする

> [!IMPORTANT]
> この記事は、シングル サインオン (SSO) クイック スタートを完了することによって作成される SSO 対応アドイン [に基になっています](sso-quickstart.md)。 この記事を読む前に、クイック スタートを完了してください。

[SSO クイック スタートは](sso-quickstart.md)、サインインしているユーザーのプロファイル情報を取得し、ドキュメントまたはメッセージに書き込む SSO 対応アドインを作成します。 この記事では、SSO クイック スタートで Yeoman ジェネレーターを使用して作成したアドインを更新し、さまざまなアクセス許可を必要とする新しい機能を追加するプロセスについて説明します。

## <a name="prerequisites"></a>前提条件

- SSO Office手順に従って作成したアドインの[1 つ。](sso-quickstart.md)

- Microsoft 365 サブスクリプションの OneDrive for Business に保存されているファイルとフォルダーが少なくともいくつかあります。

- [Node.js](https://nodejs.org) (最新 [LTS](https://nodejs.org/about/releases) バージョン)。

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>プロジェクトの内容を確認する

最初に、Yeoman ジェネレーターで以前に作成したアドイン プロジェクトを [簡単に確認します](sso-quickstart.md)。

> [!NOTE]
> この記事で **.js** ファイル拡張子を使用してスクリプト ファイルを参照する場所では、プロジェクトが TypeScript で作成されている場合は、代わりに **.ts** ファイル拡張子を想定します。

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>新しい機能を追加する

SSO クイック スタートで作成したアドインは、Microsoft Graph を使用してサインインしているユーザーのプロファイル情報を取得し、その情報をドキュメントまたはメッセージに書き込みます。 サインインしているユーザーの OneDrive for Business から上位 10 のファイルとフォルダーの名前を取得し、その情報をドキュメントまたはメッセージに書き込むアドインの機能を変更しましょう。 この新機能を有効にする場合は、Azure でアプリのアクセス許可を更新し、アドイン プロジェクト内のコードを更新する必要があります。

### <a name="update-app-permissions-in-azure"></a>Azure でアプリのアクセス許可を更新する

アドインがユーザーの OneDrive for Business のコンテンツを正常に読み取る前に、Azure のアプリ登録情報を適切なアクセス許可で更新する必要があります。 次の手順を実行して、アプリに **Files.Read.All** アクセス許可を付与し **、User.Read** アクセス許可を取り消します。これは不要です。

1. Azure Portal [に移動し](https://ms.portal.azure.com/#home)**、Microsoft 365 管理者の資格情報を使用してサインインします**。

2. [アプリの登録 **] ページに移動** します。
    > [!TIP]
    > これを行うには、Azure のホームページで [アプリの登録] タイルを選択するか、ホーム ページの検索ボックスを使用してアプリの登録を検索して **選択** します。

3. [ **アプリの登録] ページ** で、クイック スタート時に作成したアプリを選択します。
    > [!TIP]
    > アプリ **の表示** 名は、Yeoman ジェネレーターを使用してプロジェクトを作成するときに指定したアドイン名と一致します。

4. アプリの概要ページで、ページの左側にある [管理] 見出しの下にある **API** アクセス許可を選択します。

5. アクセス許可 **テーブルの User.Read** 行で、省略記号を選択し、表示されるメニューから [管理者の同意を取り消す] を選択します。

6. 表示される **プロンプトに応答** して [はい、削除] ボタンを選択します。

7. アクセス許可 **テーブルの User.Read** 行で、省略記号を選択し、表示されるメニューから [アクセス許可の削除] を選択します。

8. 表示される **プロンプトに応答** して [はい、削除] ボタンを選択します。

9. **[アクセス許可の追加]** ボタンを選択します。

10. 開いたパネルで **Microsoft Graph を選択し** 、[委任されたアクセス **許可] を選択します**。

11. [API アクセス **許可の要求] パネルで、次の操作を** 行います。

    a. [**ファイル] で****、[Files.Read.All] を選択します**。

    b. パネルの **下部にある [アクセス許可の** 追加] ボタンを選択して、これらのアクセス許可の変更を保存します。

12. [テナント **名] ボタンに管理者の同意を付与するを選択** します。

13. 表示される **プロンプトに** 応答して [はい] ボタンを選択します。

### <a name="update-code-in-the-add-in-project"></a>アドイン プロジェクトのコードを更新する

サインインしているユーザーの OneDrive for Business のコンテンツをアドインが読み取り可能にするには、次の必要があります。

- Microsoft Graph の URL、パラメーター、および必要なアクセス スコープを参照するコードを更新します。

- 作業ウィンドウ UI を定義するコードを更新して、新しい機能を正確に説明します。

- Microsoft Graph からの応答を解析し、ドキュメントまたはメッセージに書き込むコードを更新します。

次の手順では、これらの更新プログラムについて説明します。

### <a name="changes-required-for-any-type-of-add-in"></a>任意の種類のアドインに必要な変更

アドインの次の手順を実行し、Microsoft Graph の URL、パラメーター、およびアクセス スコープを変更し、作業ウィンドウ UI を更新します。 これらの手順は、アドインがターゲットとするOfficeに関係なく同じです。

1. **./.ENV** ファイル:

    a. 次 `GRAPH_URL_SEGMENT=/me` のコードに置き換える: `GRAPH_URL_SEGMENT=/me/drive/root/children`

    b. 次 `QUERY_PARAM_SEGMENT=` のコードに置き換える: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    c. 次 `SCOPE=User.Read` のコードに置き換える: `SCOPE=Files.Read.All`

2. In **./manifest.xml,** find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>` .

3. **./src/helpers/fallbackauthdialog.js** (または TypeScript プロジェクトの **./src/helpers/fallbackauthdialog.ts)** で、次のように定義されている文字列を見つけて文字列に置き換えます `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All` `requestObj` 。

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. **./src/taskpane/taskpane.html** で要素を見つけ、その要素内のテキストを更新して、アドインの新機能を `<section class="ms-firstrun-instructionstep__header">` 説明します。

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. in **./src/taskpane/taskpane.html,** find and replace both occurrences of the `Get My User Profile Information` string with the string `Read my OneDrive for Business` .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. in **./src/taskpane/taskpane.html,** find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. Microsoft Graph からの応答を解析し、ドキュメントまたはメッセージに書き込むコードを、アドインの種類に対応するセクションのガイダンスに従って更新します。

    - [Excel アドインに必要な変更 (JavaScript)](#changes-required-for-an-excel-add-in-javascript)
    - [Excel アドインに必要な変更 (TypeScript)](#changes-required-for-an-excel-add-in-typescript)
    - [Outlook アドインに必要な変更 (JavaScript)](#changes-required-for-an-outlook-add-in-javascript)
    - [Outlook アドインに必要な変更 (TypeScript)](#changes-required-for-an-outlook-add-in-typescript)
    - [PowerPoint アドインに必要な変更 (JavaScript)](#changes-required-for-a-powerpoint-add-in-javascript)
    - [PowerPoint アドインに必要な変更 (TypeScript)](#changes-required-for-a-powerpoint-add-in-typescript)
    - [Word アドインに必要な変更 (JavaScript)](#changes-required-for-a-word-add-in-javascript)
    - [Word アドインに必要な変更 (TypeScript)](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a>Excel アドインに必要な変更 (JavaScript)

アドインが JavaScript で作成された Excel アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。

1. 関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. 関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. 関数を `writeDataToExcel` 検索し、次の関数に置き換える。

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. 関数を削除 `writeDataToOutlook` します。

5. 関数を削除 `writeDataToPowerPoint` します。

6. 関数を削除 `writeDataToWord` します。

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-an-excel-add-in-typescript"></a>Excel アドインに必要な変更 (TypeScript)

アドインが TypeScript で作成された Excel アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }

    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-an-outlook-add-in-javascript"></a>Outlook アドインに必要な変更 (JavaScript)

アドインが JavaScript で作成された Outlook アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。

1. 関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. 関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. 関数を `writeDataToOutlook` 検索し、次の関数に置き換える。

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. 関数を削除 `writeDataToExcel` します。

5. 関数を削除 `writeDataToPowerPoint` します。

6. 関数を削除 `writeDataToWord` します。

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-an-outlook-add-in-typescript"></a>Outlook アドインに必要な変更 (TypeScript)

アドインが TypeScript で作成された Outlook アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }

    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a>PowerPoint アドインに必要な変更 (JavaScript)

アドインが JavaScript で作成された PowerPoint アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。

1. 関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. 関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. 関数を `writeDataToPowerPoint` 検索し、次の関数に置き換える。

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. 関数を削除 `writeDataToExcel` します。

5. 関数を削除 `writeDataToOutlook` します。

6. 関数を削除 `writeDataToWord` します。

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a>PowerPoint アドインに必要な変更 (TypeScript)

アドインが TypeScript で作成された PowerPoint アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-a-word-add-in-javascript"></a>Word アドインに必要な変更 (JavaScript)

アドインが JavaScript で作成された Word アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。

1. 関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. 関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. 関数を `writeDataToWord` 検索し、次の関数に置き換える。

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. 関数を削除 `writeDataToExcel` します。

5. 関数を削除 `writeDataToOutlook` します。

6. 関数を削除 `writeDataToPowerPoint` します。

これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。

### <a name="changes-required-for-a-word-add-in-typescript"></a>Word アドインに必要な変更 (TypeScript)

アドインが TypeScript で作成された Word アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

これらの変更を行った後、この記事の「[](#try-it-out)試してみる」セクションに進み、更新されたアドインを試してください。

## <a name="try-it-out"></a>試してみる

アドインが Excel、Word、または PowerPoint のアドインである場合は、次のセクションの手順を実行して試してください。アドインが Outlook アドインの場合は、代わりに [Outlook](#outlook) セクションの手順を完了します。

### <a name="excel-word-and-powerpoint"></a>Excel、Word、および PowerPoint

Excel、Word、または PowerPoint アドインを試すには、次の手順を実行します。

1. プロジェクトのルート フォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル Web サーバーを起動し、以前に選択した Office クライアント アプリケーションでアドインをサイドロードします。

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

    ```command&nbsp;line
    npm start
    ```

2. 前のコマンド (Excel、Word、PowerPoint など) を実行するときに開く Office クライアント アプリケーションで、アプリの [SSO](sso-quickstart.md#configure-sso) の構成中に Azure への接続に使用した Microsoft 365 管理者アカウントと同じ Microsoft 365 組織のメンバーであるユーザーとサインインしている必要があります。 これにより、SSO を正常に実行するための適切な条件が確立されます。 

3. Office クライアント アプリケーションで、[**ホーム**] タブを選択し、リボンの [**作業ウィンドウの表示**] ボタンをクリックして、アドインの作業ウィンドウを開きます。 次の画像は、Excel のこのボタンを示しています。

    ![Excel リボンで強調表示されたアドイン ボタンを示すスクリーンショット](../images/excel-quickstart-addin-3b.png)

4. 作業ウィンドウの下部にある **[OneDrive for Business** の読み取り] ボタンを選択して、SSO プロセスを開始します。

5. アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。 これは、アドインが Microsoft Graph にアクセスすることに対してテナント管理者が同意を与えていない場合、または、ユーザーが有効な Microsoft アカウント、Microsoft 365 Education または職場アカウントで Office にサインインしていない場合に発生することがあります。 ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログを示すスクリーンショット](../images/sso-permissions-request.png)

    > [!NOTE]
    > ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。

6. アドインは、サインインしているユーザーの OneDrive for Business からデータを読み取り、上位 10 のファイルとフォルダーの名前をドキュメントに書き込みます。 次の図は、Excel ワークシートに書き込まれたファイル名とフォルダー名の例を示しています。

    ![Excel ワークシートの OneDrive for Business 情報を示すスクリーンショット](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Outlook アドインを試すには、次の手順を実行します。

1. プロジェクトのルート フォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル Web サーバーを起動し、アドインをサイドロードします。 

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。 変更を行うには、管理者としてコマンド プロンプトまたはターミナルを実行する必要がある場合もあります。

    ```command&nbsp;line
    npm start
    ```

2. アプリの [SSO](sso-quickstart.md#configure-sso) の構成中に Azure への接続に使用した Microsoft 365 管理者アカウントと同じ Microsoft 365 組織のメンバーであるユーザーと Outlook にサインインしている必要があります。 これにより、SSO を正常に実行するための適切な条件が確立されます。

3. Outlook で新しいメッセージを作成します。

4. [メッセージ作成] ウィンドウで、リボンの [**作業ウィンドウの表示**] ボタンを選択して、アドインの作業ウィンドウを開きます。

    ![Outlook の [メッセージの作成] ウィンドウの [強調表示されたアドイン] リボンボタンを示すスクリーン ショット](../images/outlook-sso-ribbon-button.png)

5. 作業ウィンドウの下部にある **[OneDrive for Business** の読み取り] ボタンを選択して、SSO プロセスを開始します。

6. アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。 これは、アドインが Microsoft Graph にアクセスすることに対してテナント管理者が同意を与えていない場合、または、ユーザーが有効な Microsoft アカウント、Microsoft 365 Education または職場アカウントで Office にサインインしていない場合に発生することがあります。 ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログをのスクリーンショット](../images/sso-permissions-request.png)

    > [!NOTE]
    > ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。

7. アドインは、サインインしているユーザーの OneDrive for Business からデータを読み取り、上位 10 のファイルとフォルダーの名前を電子メール メッセージの本文に書き込みます。

    ![Outlook のメッセージ作成ウィンドウに OneDrive for Business の情報を示すスクリーンショット](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>次の手順

これで完了です。SSO クイック スタートで Yeoman ジェネレーターを使用して作成した SSO 対応アドインの機能が [正常にカスタマイズされました](sso-quickstart.md)。 Yeoman ジェネレーターが自動的に完了した SSO の構成手順、および SSO プロセスを容易にするコードの詳細については、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)
- [シングル サインオン (SSO) のクイック スタート](sso-quickstart.md)
- [シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)
- [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](../develop/troubleshoot-sso-in-office-add-ins.md)
