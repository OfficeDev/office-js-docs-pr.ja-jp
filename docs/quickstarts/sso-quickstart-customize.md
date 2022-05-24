---
title: SSO クイック スタート プロジェクトに Microsoft Graph機能を追加する
description: 作成した SSO 対応アドインに新しい Microsoft Graph機能を追加する方法について説明します。
ms.date: 05/19/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: dbcb32c14824448d2c4309df437c93d01b868288
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650630"
---
# <a name="add-microsoft-graph-functionality-to-your-sso-quick-start-project"></a>SSO クイック スタート プロジェクトに Microsoft Graph機能を追加する

> [!IMPORTANT]
> この記事は、 [シングル サインオン (SSO) クイック スタート](sso-quickstart.md)を完了して作成された SSO 対応アドインに基づいています。 この記事を読む前に、クイック スタートを完了してください。

[SSO クイック スタート](sso-quickstart.md)では、サインインしているユーザーのプロファイル情報を取得し、ドキュメントまたはメッセージに書き込む SSO 対応アドインが作成されます。 この記事では、SSO クイック スタートで Yeoman ジェネレーターで作成したアドインを更新するプロセスについて説明し、異なるアクセス許可を必要とする新機能を追加します。

## <a name="prerequisites"></a>前提条件

- [SSO クイック スタート](sso-quickstart.md)の手順に従って作成したOffice アドイン。

- Microsoft 365 サブスクリプションのOneDrive for Businessに格納されている少なくともいくつかのファイルとフォルダー。

- [Node.js](https://nodejs.org) (最新 [LTS](https://nodejs.org/about/releases) バージョン)。

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>プロジェクトの内容を確認する

まず、 [Yeoman ジェネレーター](sso-quickstart.md)を使用して以前に作成したアドイン プロジェクトの簡単なレビューを見てみましょう。

> [!NOTE]
> この記事では、 **.js** ファイル拡張子を使用してスクリプト ファイルを参照する場所で、プロジェクトが TypeScript で作成された場合は、代わりに **.ts** ファイル拡張子を想定します。

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>新しい機能を追加する

SSO クイック スタートで作成したアドインでは、Microsoft Graphを使用してサインインしているユーザーのプロファイル情報を取得し、その情報をドキュメントまたはメッセージに書き込みます。 アドインの機能を変更し、サインインしているユーザーのOneDrive for Businessから上位 10 個のファイルとフォルダーの名前を取得し、その情報をドキュメントまたはメッセージに書き込みます。 この新機能を有効にするには、Azure でアプリのアクセス許可を更新し、アドイン プロジェクト内のコードを更新する必要があります。

### <a name="update-app-permissions-in-azure"></a>Azure でアプリのアクセス許可を更新する

アドインがユーザーのOneDrive for Businessの内容を正常に読み取るには、Azure のアプリ登録情報を適切なアクセス許可で更新する必要があります。 次の手順を実行して、アプリに **Files.Read.All** アクセス許可を付与し、 **不要になった User.Read** アクセス許可を取り消します。

1. **Microsoft 365管理者の資格情報** を [使用](https://portal.azure.com)してAzure portalにサインインします。

1. **アプリの登録** ページに移動し、クイック スタート時に作成したアプリの登録を選択します。
    > [!TIP]
    > アプリの **表示名** は、Yeoman ジェネレーターを使用してプロジェクトを作成したときに指定したアドイン名と一致します。

1. [ **管理**] で [ **API アクセス許可**] を選択します。

1. アクセス許可テーブルの **User.Read** 行で省略記号を選択し、表示されるメニューから **[管理者の同意を取り消す** ] を選択します。

    :::image type="content" source="../images/app-registration-revoke-admin-consent.png" alt-text="[API アクセス許可] ページの [管理者の同意の取り消し] ボタンのスクリーンショット。":::

1. 表示されるプロンプトに応答して **、[はい、削除** ] ボタンを選択します。

1. アクセス許可テーブルの **User.Read** 行で省略記号を選択し、表示されるメニューから [ **アクセス許可の削除** ] を選択します。

    :::image type="content" source="../images/app-registration-remove-permission.png" alt-text="[API アクセス許可] ページの [アクセス許可の削除] ボタンのスクリーンショット。":::

1. 表示されるプロンプトに応答して **、[はい、削除** ] ボタンを選択します。

1. **[アクセス許可の追加]** ボタンを選択します。

1. 開いたパネルで[**Microsoft Graph**]、[**委任されたアクセス許可**] の順に選択します。

1. **[要求 API のアクセス許可**] パネルで、次の操作を行います。

    a.  [ **ファイル]** で [ **Files.Read.All**] を選択します。

    b. パネルの下部にある [ **アクセス許可の追加** ] ボタンを選択して、これらのアクセス許可の変更を保存します。

1. **[テナント名] ボタンの [管理者の同意を付与する]** を選択します。

1. 表示されるプロンプトに応答して **、[はい** ] ボタンを選択します。

### <a name="update-code-in-the-add-in-project"></a>アドイン プロジェクトのコードを更新する

アドインでサインインしているユーザーのOneDrive for Businessの内容を読み取るには、次の操作を行う必要があります。

- Microsoft Graph URL、パラメーター、および必要なアクセス スコープを参照するコードを更新します。

- 作業ウィンドウ UI を定義するコードを更新して、新しい機能を正確に記述します。

- Microsoft Graphからの応答を解析し、ドキュメントまたはメッセージに書き込むコードを更新します。

次の手順では、これらの更新プログラムについて説明します。

### <a name="changes-required-for-any-type-of-add-in"></a>任意の種類のアドインに必要な変更

アドインの次の手順を実行して、Microsoft Graph URL、パラメーター、およびアクセス スコープを変更し、作業ウィンドウ UI を更新します。 これらの手順は、アドインのターゲットとなるアプリケーションOffice関係なく同じです。

1. **./.ENV** ファイル:

    a.  `GRAPH_URL_SEGMENT=/me` を `GRAPH_URL_SEGMENT=/me/drive/root/children` に置き換え

    b. `QUERY_PARAM_SEGMENT=` を `QUERY_PARAM_SEGMENT=?$select=name&$top=10` に置き換え

    c. `SCOPE=User.Read` を `SCOPE=Files.Read.All` に置き換え

1. **./manifest.xml** で、ファイルの末尾付近の行`<Scope>User.Read</Scope>`を検索し、行`<Scope>Files.Read.All</Scope>`に置き換えます。

1. **./src/helpers/fallbackauthdialog.js** (または TypeScript プロジェクトの **./src/helpers/fallbackauthdialog.ts**) で、文字列`https://graph.microsoft.com/User.Read`を検索し、次のように定義されている`requestObj`文字列`https://graph.microsoft.com/Files.Read.All`に置き換えます。

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

1. **./src/taskpane/taskpane.html** で、要素を検索し、その要素`<section class="ms-firstrun-instructionstep__header">`内のテキストを更新してアドインの新機能を記述します。

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

1. **./src/taskpane/taskpane.html** で、文字列`Get My User Profile Information`の出現箇所の両方を検索し、それを `Read my OneDrive for Business`.

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

1. **./src/taskpane/taskpane.html** で文字列`Your user profile information will be displayed in the document.`を見つけて、それを `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

1. アドインの種類に対応するセクションのガイダンスに従って、Microsoft Graphからの応答を解析し、ドキュメントまたはメッセージに書き込むコードを更新します。

    - [Office アドインに必要な変更 (JavaScript)](#changes-required-for-an-office-add-in-javascript)
    - [Office アドインに必要な変更 (TypeScript)](#changes-required-for-an-office-add-in-typescript)

### <a name="changes-required-for-an-office-add-in-javascript"></a>Office アドインに必要な変更 (JavaScript)

生成されたOfficeアドインで JavaScript を使用している場合は、**./src/helpers/documentHelper.js** で次の変更を行います。

1. 関数を `filterUserProfileInfo` 見つけて、次の関数に置き換えます。

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

1. `filterUserProfileInfo`検索して置き換えます`filterOneDriveInfo`。 置換するインスタンスは 4 つ必要です。

1. 変更を保存します。

これらの変更を行った後は、この記事の「試してみる」セクションに進 [み](#try-it-out) 、更新されたアドインを試してください。

### <a name="changes-required-for-an-office-add-in-typescript"></a>Office アドインに必要な変更 (TypeScript)

生成されたOfficeアドインで TypeScript が使用されている場合は、**./src/taskpane/taskpane.ts** を開きます。

1. この関数を`writeDataToOfficeDocument`見つけて、アドインで使用Officeホストに応じて次のコードに置き換えます (Excel、Outlook、Word、またはPowerPoint)

#### <a name="excel-code"></a>Excel コード

```typescript
  export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[][];

    // Get just the filenames from results
    data = result["value"].map((item) => {
      return [item.name];
    });

    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

#### <a name="outlook-code"></a>Outlookコード

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  // Get just the filenames from results.
  const data: string[] = result["value"].map((item) => {
    return item.name;
  });

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "</br>";
  }
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
```

#### <a name="word-code"></a>Word コード

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function (context) {
    // Get just the filenames from results.
    const data: string[] = result["value"].map((item) => {
      return item.name;
    });

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

#### <a name="powerpoint-code"></a>PowerPointコード

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  // Get just the filenames from results.
  const data: string[] = result["value"].map((item) => {
    return item.name;
  });
  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

## <a name="try-it-out"></a>試してみる

アドインがExcel、Word、またはPowerPoint アドインの場合は、次のセクションの手順を実行して試してください。アドインが Outlook アドインの場合は、代わりに [[Outlook](#outlook)] セクションの手順を実行します。

### <a name="excel-word-and-powerpoint"></a>Excel、Word、および PowerPoint

Excel、Word、または PowerPoint アドインを試すには、次の手順を実行します。

1. プロジェクトのルート フォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル Web サーバーを起動し、以前に選択したOfficeクライアント アプリケーションにアドインをサイドロードします。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. 前のコマンド (Excel、Word、PowerPoint など) を実行したときに開くOffice クライアント アプリケーションで、[SSO の構成](sso-quickstart.md#configure-sso)中に Azure への接続に使用したMicrosoft 365管理者アカウントと同じMicrosoft 365組織のメンバーであるユーザーでサインインしていることを確認します。 アプリの場合。 これにより、SSO を正常に実行するための適切な条件が確立されます。 

3. Office クライアント アプリケーションで、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。次の画像は、Excel のこのボタンを示しています。

    ![リボンの強調表示されたアドイン ボタンExcel示すスクリーンショット。](../images/excel-quickstart-addin-3b.png)

4. 作業ウィンドウの下部にある [**OneDrive for Businessの読み取り**] ボタンを選択して、SSO プロセスを開始します。

5. アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。 これは、アドインが Microsoft Graph にアクセスすることに対してテナント管理者が同意を与えていない場合、または、ユーザーが有効な Microsoft アカウント、Microsoft 365 Education または職場アカウントで Office にサインインしていない場合に発生することがあります。 ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログを示すスクリーンショット。](../images/sso-permissions-request.png)

    > [!NOTE]
    > ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。

6. アドインは、サインインしているユーザーのOneDrive for Businessからデータを読み取り、上位 10 個のファイルとフォルダーの名前をドキュメントに書き込みます。 次の図は、Excel ワークシートに書き込まれたファイル名とフォルダー名の例を示しています。

    ![ワークシートのOneDrive for Business情報Excel示すスクリーンショット。](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Outlook アドインを試すには、次の手順を実行します。

1. プロジェクトのルート フォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル Web サーバーを起動し、アドインをサイドロードします。 

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. アプリの [SSO の構成](sso-quickstart.md#configure-sso)中に、Azure への接続に使用したMicrosoft 365管理者アカウントと同じMicrosoft 365組織のメンバーであるユーザーと共にOutlookにサインインしていることを確認します。 これにより、SSO を正常に実行するための適切な条件が確立されます。

3. Outlook で新しいメッセージを作成します。

4. [メッセージ作成] ウィンドウで、リボンの [**作業ウィンドウの表示**] ボタンを選択して、アドインの作業ウィンドウを開きます。

    ![Outlook の [メッセージの作成] ウィンドウの [強調表示されたアドイン] リボン ボタンを示すスクリーン ショット。](../images/outlook-sso-ribbon-button.png)

5. 作業ウィンドウの下部にある [**OneDrive for Businessの読み取り**] ボタンを選択して、SSO プロセスを開始します。

6. アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。 これは、アドインが Microsoft Graph にアクセスすることに対してテナント管理者が同意を与えていない場合、または、ユーザーが有効な Microsoft アカウント、Microsoft 365 Education または職場アカウントで Office にサインインしていない場合に発生することがあります。 ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログのスクリーンショット。](../images/sso-permissions-request.png)

    > [!NOTE]
    > ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。

7. アドインは、サインインしているユーザーのOneDrive for Businessからデータを読み取り、上位 10 個のファイルとフォルダーの名前を電子メール メッセージの本文に書き込みます。

    ![Outlookメッセージの作成ウィンドウにOneDrive for Business情報を示すスクリーンショット。](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>次の手順

おめでとうございます。SSO [クイック スタート](sso-quickstart.md)で Yeoman ジェネレーターを使用して作成した SSO 対応アドインの機能が正常にカスタマイズされました。 Yeoman ジェネレーターが自動的に完了した SSO の構成手順、および SSO プロセスを容易にするコードの詳細については、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)
- [シングル サインオン (SSO) のクイック スタート](sso-quickstart.md)
- [シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)
- [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](../develop/troubleshoot-sso-in-office-add-ins.md)
