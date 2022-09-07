---
title: シングル サインオンを使用する Node.js Office アドインを作成する
description: Office シングル サインオンを使用するNode.js ベースのアドインを作成する方法について説明します。
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e7ded29d9d2f021516348e2edbe847b6447e006
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616050"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する Node.js Office アドインを作成する

ユーザーは、このサインイン プロセスを利用してユーザーを承認する Office および Office Web アドインにサインインできます。こうして承認されたユーザーは、アドインと Microsoft Graph への 2 度目のサインオンの必要がなくなります。概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。

この記事では、アドインでシングル サインオン (SSO) を有効にするプロセスについて説明します。 作成するサンプル アドインには 2 つの部分があります。Microsoft Excel に読み込む作業ウィンドウと、作業ウィンドウの Microsoft Graph への呼び出しを処理する中間層サーバー。 中間層サーバーは、Node.jsと Express で構築され、単一の REST API を公開します。これは、 `/getuserfilenames`ユーザーの OneDrive フォルダー内の最初の 10 個のファイル名の一覧を返します。 作業ウィンドウでは、このメソッドを `getAccessToken()` 使用して、サインインしているユーザーのアクセス トークンを中間層サーバーに取得します。 中間層サーバーでは、On-Behalf-Of フロー (OBO) を使用して、Microsoft Graph にアクセスできる新しいサーバーとアクセス トークンを交換します。 このパターンを拡張して、任意の Microsoft Graph データにアクセスできます。 作業ウィンドウは、Microsoft Graph サービスが必要な場合、常に中間層 REST API を呼び出します (アクセス トークンを渡します)。 中間層では、OBO を使用して取得したトークンを使用して Microsoft Graph サービスを呼び出し、結果を作業ウィンドウに返します。

この記事は、Node.jsと Express を使用するアドインで動作します。 ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)

- [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

- コード エディター - Visual Studio Code をお勧めします

- Microsoft 365 サブスクリプションのOneDrive for Businessに格納されている少なくともいくつかのファイルとフォルダー

- [IdentityAPI 1.3 要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)をサポートする Microsoft 365 のビルド。 90 日間の再生可能なMicrosoft 365 E5開発者サブスクリプションを提供する無料の開発者[サンドボックス](https://developer.microsoft.com/microsoft-365/dev-program#Subscription)を入手できます。 開発者サンドボックスには、この記事の後の手順でアプリの登録に使用できる Microsoft Azure サブスクリプションが含まれています。 必要に応じて、アプリの登録に別の Microsoft Azure サブスクリプションを使用できます。 [Microsoft Azure](https://account.windowsazure.com/SignUp) で試用版サブスクリプションを取得します。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

1. 「[Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)」にあるリポジトリを複製するかダウンロードします。

   > [!NOTE]
   > サンプルには 2 つのバージョンがあります。
   >
   > - **Begin** フォルダーはスターター プロジェクトです。 SSO や承認に直接関連しない UI などの側面は、既に完了しています。 この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。
   > - **Complete** フォルダーには、この記事のすべてのコーディング手順が完了した同じサンプルが含まれています。 完成したバージョンを使用するには、この記事の手順に従いますが、"Begin" を "Complete" に置き換え、セクション「 **クライアント側のコード** 化」と「 **中間層サーバー側のコード** 化」をスキップします。

1. **[開始]** フォルダーでコマンド プロンプトを開きます。

1. コンソールで `npm install` を入力して、package.json ファイルに項目化されているすべての依存関係をインストールします。

1. コマンド`npm run install-dev-certs`を実行します。 証明書をインストールするプロンプトに対して **はい** を選択します。

## <a name="register-the-add-in-with-microsoft-identity-platform"></a>アドインをMicrosoft ID プラットフォームに登録する

中間層サーバーを表すアプリ登録を Azure で作成する必要があります。 これにより、JavaScript でクライアント コードに適切なアクセス トークンを発行できるように、認証のサポートが有効になります。 この登録では、クライアントでの SSO と、Microsoft 認証ライブラリ (MSAL) を使用したフォールバック認証の両方がサポートされます。

1. アプリを登録するには、[Azure portal - アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908) ページに移動してアプリを登録します。

1. **_管理者_** 資格情報を使用して Microsoft 365 テナントにサインインします。 たとえば、MyName@contoso.onmicrosoft.com です。

1. **[新規登録]** を選択します。 **[アプリケーションを登録]** ページで、次のように値を設定します。

   - `Office-Add-in-NodeJS-SSO` に **[名前]** を設定します。
   - **サポートされているアカウントの種類** を **、任意の組織ディレクトリ (任意の Azure AD ディレクトリ - マルチテナント) および個人用 Microsoft アカウント (Skype、Xbox など) のアカウント** に設定します。
   - [**リダイレクト URI**] セクションで、リダイレクト URI の値`https://localhost:44355/dialog.html`が . の **単一ページ アプリケーション (SPA)** にプラットフォームを設定します。
   - **[登録]** を選択します。

   > [!NOTE]
   > SPA アプリケーションの種類は、クライアントがフォールバック認証に MSAL を使用する場合にのみ使用されます。

1. **Office-Add-in-NodeJS-SSO** ページで、**アプリケーション (クライアント) ID** と **ディレクトリ (テナント) ID** の値をコピーして保存します。 以降の手順では、それらの両方を使用します。

   > [!NOTE]
   > この **アプリケーション (クライアント) ID** は、Office クライアント アプリケーション (PowerPoint、Word、Excel など) などの他のアプリケーションがアプリケーションへの承認されたアクセスを求める場合の "対象ユーザー" の値です。 また、Microsoft Graph への承認されたアクセスを求めるアプリケーションの "クライアント ID" でもあります。

1. 左端のサイドバーで、[**管理**] で [**認証**] を選択します。 [ **暗黙的な付与とハイブリッド フロー** ] セクションで、 **Access トークン** と **ID トークン** の両方のチェック ボックスをオンにします。 このサンプルでは、SSO を使用できない場合にフォールバック認証に Microsoft 認証ライブラリ (MSAL) を使用します。

1. **[保存]** を選択します。

1. [ **管理**] で、[ **証明書&シークレット** ] を選択し、[ **新しいクライアント シークレット**] を選択します。 [**説明**] に値を入力してから、[**有効期限**] の適切なオプションを選択し、[**追加**] を選択します。

   Web アプリケーションは、クライアント シークレット **値** を使用して、トークンを要求するときにその ID を証明します。 _後の手順で使用するためにこの値を記録します。この値は 1 回だけ表示されます。_

1. 左端のサイドバーで、[**管理**] で [**API の公開**] を選択します。 [ **設定** ] リンクを選択します。 これにより、"api://$App ID GUID$" という形式でアプリケーション ID URI が生成されます。ここで、$App ID GUID$ は **アプリケーション (クライアント) ID です**。

1. 生成された ID で、二重スラッシュと GUID の間に挿入 `localhost:44355/` (末尾にスラッシュ "/" が追加されていることに注意してください)。 完了したら、ID 全体にフォーム `api://localhost:44355/$App ID GUID$`が含まれている必要があります 。たとえば、次のようになります `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。 次に、**[保存]** を選択します。

1. **[Scope の追加]** ボタンをクリックします。 開いたパネルで、**[スコープ名]** として `access_as_user` を入力します。

1. **[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。

1. 管理者とユーザーの同意プロンプトを構成するためのフィールドに、Office クライアント アプリケーションが現在のユーザーと同じ権限でアドインの Web API を使用できるようにするスコープに適 `access_as_user` した値を入力します。 提案:

   - **管理同意表示名**: Office はユーザーとして機能できます。
   - **管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。
   - **ユーザーの同意表示名**: Office は、ユーザーの役割を果たすことができます。
   - **ユーザーの同意の説明**: Office が、自分と同じ権限を持つアドインの Web API を呼び出すようにします。

1. **[状態]** が **[有効]** に設定されていることを確認してください。

1. **[スコープの追加]** を選択します。

   > [!NOTE]
   > テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。

1. [ **承認されたクライアント アプリケーション** ] セクションで、[ **クライアント アプリケーションの追加]** ボタンを選択し、開いたパネルで [クライアント ID] を `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`[ **承認されたスコープ** ] チェック ボックスをオンにします `api://localhost:44355/$app-id-guid$/access_as_user`。

1. **[アプリケーションの追加]** を選択します。

   > [!NOTE]
   > この ID は `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` 、すべての Microsoft Office アプリケーション エンドポイントを事前に承認します。 また、Windows および Mac 上の Office で Microsoft アカウント (MSA) をサポートする場合にも必要です。 または、何らかの理由で一部のプラットフォームで Office への承認を拒否する場合は、次の ID の適切なサブセットを入力することもできます。 承認を保留するプラットフォームの ID は残しておきます。 これらのプラットフォーム上のアドインのユーザーは、Web API を呼び出すことはできませんが、アドイン内の他の機能は引き続き機能します。
   >
   > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
   > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
   > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. 左端のサイドバーで、[**管理**] で **[API アクセス許可**] を選択し、[**アクセス許可の追加]** を選択します。 開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。

1. アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。 以下を選択します。 アドイン自体で実際に必要なのは 1 つ目だけです。ただし、 `profile` Office アプリケーションが中間層サーバーにアクセスするためにユーザー ID を持つアクセス トークンを取得するには、アクセス許可と `openid` アクセス許可が必要です。

   - **Files.Read**
   - **profile**
   - **openid**

   > [!NOTE]
   > `User.Read` アクセス許可は既定でリストされています。 必要のないアクセス許可を要求しないことをお勧めします。そのため、アドインで実際に必要ない場合は、このアクセス許可のチェック ボックスをオフにすることをお勧めします。

1. 表示される各アクセス許可のチェック ボックスをオンにします。 アドインに必要なアクセス許可を選択したら、パネルの下部にある **[アクセス許可を追加する]** ボタンをクリックします。

1. 同じページで、**[[テナント名]に管理者の同意を与える]** ボタンを選択し、表示される確認に対して **[はい]** を選択します。

## <a name="configure-the-add-in"></a>アドインを構成する

1. コード エディターで複製プロジェクトの`\Begin`フォルダーを開きます。

1. ファイルを `.ENV` 開き、 **Office-Add-in-NodeJS-SSO** アプリの登録から前にコピーした値を使用します。 次のように値を設定します。

   | 名前              | 値                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | アプリ登録の概要ページからの **アプリケーション (クライアント) ID**。 |
   | **CLIENT_SECRET** | **[証明書] & [シークレット]** ページから保存された **クライアント** シークレット。       |
   | **DIRECTORY_ID**  | アプリ登録の概要ページからの **ディレクトリ (テナント) ID**。   |

   値は引用符で囲ま **ない** でください。 完了すると、ファイルは以下のようになります。

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. マークアップ内の _両方の場所にある_ プレースホルダー "$app-id-guid$" を **、Office-Add-in-NodeJS-SSO** アプリ登録の作成時にコピーした **アプリケーション ID** に置き換えます。 "$" 記号は ID の一部ではないため、含めないでください。 これは、. のCLIENT_IDに使用した ID と同じです。ENV ファイル。

   > [!NOTE]
   > 値は **\<Resource\>** 、アドインを登録したときに設定した **アプリケーション ID URI** です。 この **\<Scopes\>** セクションは、アドインが AppSource から販売されている場合にのみ、同意ダイアログ ボックスを生成するために使用されます。

1. `\public\javascripts\fallback-msal\authConfig.js` ファイルを開きます。 プレースホルダー "$app-id-guid$" を、前に作成した **Office-Add-in-NodeJS-SSO** アプリ登録から保存したアプリケーション ID に置き換えます。

1. 変更をファイルに保存します。

## <a name="code-the-client-side"></a>クライアント側のコーディング

### <a name="create-client-request-and-response-handler"></a>クライアント要求と応答ハンドラーを作成する

1. コード エディターで、`public\javascripts\ssoAuthES6.js`ファイルを開きます。 Internet Explorer 11 でも Promise がサポートされることを保証するコードと、アドインの唯一のボタンにハンドラーを割り当てるための`Office.onReady`呼び出しが既にあります。

   > [!NOTE]
   > 名前が示すように、ssoAuthES6.js は JavaScript ES6 構文を使用します。これは、これは、`async`と`await`の使用こそが SSO API の本質的なシンプルさを最もよく示すためです。 localhost サーバーが起動されると、このファイルは ES5 構文に変換され、サンプルで Internet Explorer 11 がサポートされます。

    サンプル コードの重要な部分は、クライアント要求です。 クライアント要求は、中間層サーバーで REST API を呼び出すための要求に関する情報を追跡するオブジェクトです。 クライアント要求の状態は、次のシナリオで追跡または更新する必要があるため、必要です。

    - SSO が失敗し、フォールバック認証が必要です。 アクセス トークンは、ポップアップ ダイアログ ボックスで MSAL を使用して取得されます。 目標は、このシナリオでは失敗せず、代替認証アプローチに正常にフォールバックすることです。

    クライアント要求オブジェクトは、次のデータを追跡します。

    - `authSSO` - SSO を使用する場合は true、それ以外の場合は false。
    - `verb` - GET や POST などの REST API 動詞。
    - `accessToken`- ASP.NET Core サーバーへのアクセス トークン。
    - `url`- ASP.NET Core サーバーで呼び出す REST API の URL。
    - `callbackRESTApiHandler` - REST API 呼び出しの結果を渡す関数。
    - `callbackFunction` - 準備ができたらクライアント要求を渡す関数。

1. クライアント要求オブジェクトを初期化するには、関数で次の `createRequest` コードに置き換えます `TODO 1` 。

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. `TODO 2` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - SSO が使用されているかどうかを確認します。 SSO のアクセス トークンを取得する方法は、フォールバック認証の場合とは異なります。
    - SSO がアクセス トークンを返す場合は、関数を呼び出します `callbackfunction` 。 フォールバック認証では、ユーザーが MSAL 経由でサインインした後、最終的にコールバック関数を呼び `dialogFallback`出します。

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. `getFileNameList` 関数で、`TODO 3` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - この関数 `getFileNameList` は、ユーザーが作業ウィンドウの **[OneDrive ファイル名の取得** ] ボタンを選択したときに呼び出されます。
    - REST API の URL など、呼び出しに関する情報を追跡するクライアント要求を作成します。
    - REST API から結果が返されると、関数に `handleGetFileNameResponse` 渡されます。 このコールバックはパラメーター `createRequest` として渡され、追跡されます `clientRequest.callbackRESTApiHandler`。
    - 次の手順を実行し、REST API を呼び出 `callWebServer` すために、クライアント要求を使用してコードが呼び出されます。

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. `handleGetFileNameResponse` 関数で、`TODO 4` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - このコードは、ドキュメントにファイル名を書き込む応答 (ファイル名の一覧を含む) `writeFileNamesToOfficeDocument` を渡します。
    - コードはエラーをチェックします。 ファイル名が書き込まれている場合は成功メッセージが表示され、それ以外の場合はエラーが表示されます。

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. `handleSSOErrors` 関数で、`TODO 5` を次のコードに置き換えます。 これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. `TODO 6`を以下のコードに置き換えます。 これらのエラーの詳細については、「 [Office アドインでの SSO のトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。処理できないエラーについては、 `true` 呼び出し元に返されます。 これは、呼び出し元がフォールバック認証として MSAL を使用するように切り替える必要があることを示します。

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### <a name="call-the-rest-api-on-the-middle-tier-server"></a>中間層サーバーで REST API を呼び出す

1. `callWebServer` 関数で、`TODO 7` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - 実際の AJAX 呼び出しは、関数によって `ajaxCallToRESTApi` 行われます。
    - 中間層サーバーから現在のトークンの有効期限が切れたことを示すエラーが返された場合、この関数は新しいアクセス トークンの取得を試みます。
    - AJAX 呼び出しを正常に完了できない場合は、 `switchToFallbackAuth` Office SSO の代わりに MSAL 認証を使用するように呼び出されます。

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. `TODO 8` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - サーバーは、有効期限が切れたトークンを識別すると、"TokenExpiredError" 型のエラーを返します。
    - try...catch は Office.auth.getAccessToken を呼び出して、新しい有効期限で更新されたトークンを取得します。
    - コードは、サーバー API の呼び出しを再試行します。

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. `TODO 9` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - **Microsoft Graph** エラーの場合は、作業ウィンドウにメッセージを表示します。
    - その他のすべてのメッセージについては、作業ウィンドウにメッセージを表示します。

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

フォールバック認証では、MSAL ライブラリを使用してユーザーにサインインします。 アドイン自体は SPA であり、SPA アプリの登録を使用して中間層サーバーにアクセスします。

1. `switchToFallbackAuth` 関数で、`TODO 10` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - グローバル `authSSO` を false に設定し、認証に MSAL を使用する新しいクライアント要求を作成します。新しい要求には、中間層サーバーへの MSAL アクセス トークンがあります。
    - 要求が作成されると、引き続き中間層サーバーの呼び出 `callWebServer` しが正常に試行されます。

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## <a name="code-the-middle-tier-server"></a>中間層サーバーをコーディングする

中間層サーバーは、クライアントが呼び出す REST API を提供します。 たとえば、REST API `/getuserfilenames` は、ユーザーの OneDrive フォルダーからファイル名の一覧を取得します。 各 REST API 呼び出しでは、正しいクライアントがデータにアクセスしていることを確認するために、クライアントによるアクセス トークンが必要です。 アクセス トークンは、On-Behalf-Of フロー (OBO) を介して Microsoft Graph トークンと交換されます。 新しい Microsoft Graph トークンは、後続の API 呼び出しのために MSAL ライブラリによってキャッシュされます。 中間層サーバーの外部に送信されることはありません。 詳細については、「[中間層アクセス トークン要求](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)」を参照してください。

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>ルートを作成し、On-Behalf-Of フローを実装する

1. ファイル `routes\getFilesRoute.js` を開き、次のコードに置き換えます `TODO 11` 。 このコードについては、以下の点に注意してください。

    - 呼び出します `authHelper.validateJwt`。 これにより、アクセス トークンが有効になり、改ざんされていないことを確認できます。
    - 詳細については、「 [トークンの検証](/azure/active-directory/develop/access-tokens#validating-tokens)」を参照してください。

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. `TODO 12` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - 必要な最小限のスコープのみを要求します (例: `files.read`.
    - MSAL を `authHelper` 使用して、次の呼び出しで OBO フローを実行します `acquireTokenOnBehalfOf`。

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. `TODO 13` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - Microsoft Graph API呼び出しの URL を作成し、関数を使用して呼び出しを`getGraphData`行います。
    - HTTP 500 応答と詳細を送信してエラーを返します。
    - 成功すると、ファイル名リストを含む JSON がクライアントに返されます。

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    ```

1. `TODO 14` を次のコードに置き換えます。 このコードでは、クライアントが新しいトークンを要求して再度呼び出すことができるため、トークンの有効期限が切れているかどうかを具体的に確認します。

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
   }
   ```

このサンプルでは、MSAL によるフォールバック認証と Office による SSO 認証の両方を処理する必要があります。 サンプルは最初に SSO を試し `authSSO` 、サンプルで SSO を使用しているかフォールバック認証に切り替えた場合は、ファイルの先頭にあるブール値が追跡されます。

## <a name="run-the-project"></a>プロジェクトを実行する

1. 結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。

1. `\Begin`フォルダーのルートでコマンド プロンプトを開きます。

1. コマンド `npm install` を実行して、すべてのパッケージの依存関係をインストールします。

1. コマンド `npm start` を実行して中間層サーバーを起動します。

1. アドインを Office アプリケーション (Excel、Word、または PowerPoint) にサイドロードして、テストをする必要があります。 手順はプラットフォームによって異なります。 「[テスト用に Office アドインをサイドロードする](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)」に手順へのリンクがあります。

1. Office アプリケーションの **[ホーム]** リボンで **[アドインの表示]** ボタン (**SSO Node.js** グループ内) を選択して、作業ウィンドウ アドインを開きます。

1. **[OneDrive ファイル名の取得]** ボタンをクリックします。 Microsoft 365 Educationまたは職場のアカウント、または Microsoft アカウントで Office にログインしていて、SSO が想定どおりに動作している場合は、OneDrive for Businessの最初の 10 個のファイル名とフォルダー名がドキュメントに挿入されます。 (初回には 15 秒かかる場合があります)。ログインしていない場合、または SSO をサポートしていないシナリオの場合、または SSO が何らかの理由で機能していない場合は、サインインを求めるメッセージが表示されます。 サインインすると、ファイル名とフォルダー名が表示されます。

> [!NOTE]
> 以前に別の ID で Office にサインインしており、その時点で開いていた一部の Office アプリケーションがまだ開いている場合、Office が ID を変更したかのように見えても、確実に ID を変更できていない場合があります。 これが発生すると、Microsoft Graph の呼び出しが失敗するか、以前の ID のデータが返される場合があります。 これを防ぐには、必ず _他のすべての Office アプリケーションを閉じて_ から、**[OneDrive ファイル名の取得]** を押してください。

## <a name="security-notes"></a>セキュリティ に関する注意事項

- ルートでは`/getuserfilenames``getFilesroute.js`、リテラル文字列を使用して Microsoft Graph の呼び出しを作成します。 文字列の一部がユーザー入力から取得されるように呼び出しを変更する場合は、応答ヘッダーインジェクション攻撃で使用できないように入力をサニタイズします。

- `app.js`次のコンテンツ セキュリティ ポリシーは、スクリプト用に用意されています。 アドインのセキュリティ ニーズに応じて、追加の制限を指定することもできます。

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

[Microsoft ID プラットフォームのドキュメント](/azure/active-directory/develop/)では、常にセキュリティのベスト プラクティスに従ってください。
