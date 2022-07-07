---
title: SSO を使用してサインインしているユーザーの ID を取得する
description: getAccessToken API を呼び出して、サインインしているユーザーに関する名前、電子メール、追加情報を含む ID トークンを取得します。
ms.date: 02/16/2022
localization_priority: Normal
ms.openlocfilehash: 5416c469a15d7eda9333f511c61e2cff1a901018
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660068"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>SSO を使用してサインインしているユーザーの ID を取得する

API を `getAccessToken` 使用して、Office にサインインしている現在のユーザーの ID を含むアクセス トークンを取得します。 アクセス トークンは、サインインしているユーザーに関する ID 要求 (名前や電子メールなど) が含まれているため、ID トークンでもあります。 ID トークンを使用して、独自の Web サービスを呼び出すときにユーザーを識別することもできます。 呼び出 `getAccessToken` すには、Office アドインで SSO を使用するように Office アドインを構成する必要があります。

この記事では、ID トークンを取得し、ユーザーの名前、電子メール、一意の ID を作業ウィンドウに表示する Office アドインを作成します。

> [!NOTE]
> Office と API の `getAccessToken` SSO は、すべてのシナリオでは機能しません。 SSO が利用できない場合は、常にフォールバック ダイアログを実装してユーザーにサインインします。 詳細については、「 [Office ダイアログ API を使用した認証と承認」を](auth-with-office-dialog-api.md)参照してください。

## <a name="create-an-app-registration"></a>アプリの登録を作成する

Office で SSO を使用するには、Microsoft ID プラットフォームが Office アドインとそのユーザーに認証および承認サービスを提供できるように、Azure portalにアプリ登録を作成する必要があります。

1. アプリを登録するには、[Azure portal - アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908) ページに移動します。

1. **_管理者_** 資格情報を使用して Microsoft 365 テナントにサインインします。 たとえば、MyName@contoso.onmicrosoft.com です。

1. **[新規登録]** を選択します。 **[アプリケーションを登録]** ページで、次のように値を設定します。

   - `Office-Add-in-SSO` に **[名前]** を設定します。
   - **[サポートされているアカウントの種類]** を **[任意の組織のディレクトリ内のアカウントと個人用の Microsoft アカウント (例: Skype、 Xbox、Outlook.com)]** に設定します。
   - アプリケーションの種類を **Web** に設定し、[ **リダイレクト URI] を [リダイレクト URI]** に `https://localhost:[port]/dialog.html`設定します。 Web アプリケーションの正しいポート番号に置き換えます `[port]` 。 office を使用してアドインを作成した場合、通常、ポート番号は 3000 であり、package.json ファイルにあります。 Visual Studio 2019 でアドインを作成した場合、ポートは Web プロジェクトの **SSL URL** プロパティにあります。
   - **[登録]** を選択します。

1. **Office アドイン SSO ページで**、**アプリケーション (クライアント) ID** と **ディレクトリ (テナント)** ID の値をコピーして保存します。 以降の手順では、それらの両方を使用します。

   > [!NOTE]
   > この **アプリケーション (クライアント) ID** は、Office クライアント アプリケーション (PowerPoint、Word、Excel など) などの他のアプリケーションがアプリケーションへの承認されたアクセスを求める場合の "対象ユーザー" の値です。 また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. **[管理]** の下の **[認証]** を選択します。 [ **暗黙的な付与** ] セクションで、 **アクセス トークン** と **ID トークン** の両方のチェック ボックスをオンにします。

1. フォームの最上部で **[保存]** を選択します。

1. **[管理]** の下の **[API の公開]** を選択します。 [ **設定** ] リンクを選択します。 これにより、フォーム`api://[app-id-guid]``[app-id-guid]`にアプリケーション ID URI (**アプリケーション (クライアント) ID)** が生成されます。

1. 生成された ID で、二重スラッシュと GUID の間に挿入 `localhost:[port]/` (末尾にスラッシュ "/" が追加されていることに注意してください)。 Web アプリケーションの正しいポート番号に置き換えます `[port]` 。 office を使用してアドインを作成した場合、通常、ポート番号は 3000 であり、package.json ファイルにあります。 Visual Studio 2019 でアドインを作成した場合、ポートは Web プロジェクトの **SSL URL** プロパティにあります。
   完了したら、ID 全体にフォーム `api://localhost:[port]/[app-id-guid]`が含まれている必要があります 。たとえば、次のようになります `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

1. **[Scope の追加]** ボタンをクリックします。 開いたパネルで、名前として **\<Scope\>** 入力`access_as_user`します。

1. **[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。

1. 管理者とユーザーの同意プロンプトを構成するためのフィールドに、Office クライアント アプリケーションが現在のユーザーと同じ権限でアドインの Web API を使用できるようにするスコープに適 `access_as_user` した値を入力します。 提案:

   - **管理同意表示名**: Office はユーザーとして機能できます。
   - **管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。
   - **ユーザーの同意表示名**: Office は、ユーザーの役割を果たすことができます。
   - **ユーザーの同意の説明**: Office が、自分と同じ権限を持つアドインの Web API を呼び出すようにします。

1. **[状態]** が **[有効]** に設定されていることを確認してください。

1. **[スコープの追加]** を選択します。

   > [!NOTE]
   > テキスト フィールドのすぐ下に表示される名前の **\<Scope\>** ドメイン部分は、前に設定したアプリケーション ID URI と自動的に一致し、`/access_as_user`末尾に追加されます。たとえば、 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`

1. [ **承認済みクライアント アプリケーション** ] セクションで、次の ID を入力して、すべての Microsoft Office アプリケーション エンドポイントを事前に承認します。

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (すべての Microsoft Office アプリケーション エンドポイント)

    > [!NOTE]
    > ID は `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` 、次のすべてのプラットフォームで Office を事前に承認します。 または、何らかの理由で一部のプラットフォームで Office への承認を拒否する場合は、次の ID の適切なサブセットを入力することもできます。 承認を保留するプラットフォームの ID は残しておきます。 これらのプラットフォーム上のアドインのユーザーは、Web API を呼び出すことはできませんが、アドイン内の他の機能は引き続き機能します。
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. [ **クライアント アプリケーションの追加] ボタンを** 選択し、開いたパネルで [アプリケーション (クライアント) ID] に設定 `[app-id-guid]` し `api://localhost:44355/[app-id-guid]/access_as_user`、[.

1. **[アプリケーションの追加]** を選択します。

1. **[管理]** の下の **[API アクセス許可]** を選択し、**[アクセス許可の追加]** を選択します。 開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。

1. アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。 **プロファイル** のアクセス許可を検索して選択します。 `profile` Office アプリケーションがアドイン Web アプリケーションに対するトークンを取得するには、このアクセス許可が必要です。

   - profile

   > [!NOTE]
   > `User.Read` アクセス許可は既定でリストされています。 必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。

1. パネル下部にある **[アクセス許可の追加]** ボタンを選択します。

1. 同じページで、[管理者の **同意\<tenant-name\>の付与**] ボタンを選択し、表示される確認のために **[はい**] を選択します。

## <a name="create-the-office-add-in"></a>Office アドインを作成する

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Visual Studio 2019 を起動し、 **新しいプロジェクトの作成を** 選択します。
1. **Excel Web アドイン** プロジェクト テンプレートを検索して選択します。 **[次へ]** を選択します。 注: SSO は任意の Office アプリケーションで動作しますが、この記事では Excel で動作します。
1. **sso-display-user-info** などのプロジェクト名を入力し、[作成] を選択 **します**。 他のフィールドは既定値のままにできます。
1. [ **アドインの種類の選択** ] ダイアログ ボックスで、[ **Excel に新しい機能を追加** する] を選択し、[完了] を選択 **します**。

プロジェクトが作成され、ソリューションに 2 つのプロジェクトが含まれます。

- **sso-display-user-info**: アドインを Excel にサイドロードするためのマニフェストと詳細が含まれます。
- **sso-display-user-infoWeb**: アドインの Web ページをホストする ASP.NET プロジェクト。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

[開発環境を設定](../overview/set-up-your-dev-environment.md)していることを確認します。

1. 次のコマンドを入力してプロジェクトを作成します。

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

プロジェクトは、 **sso-display-user-info** という名前の新しいフォルダーに作成されます。

---

## <a name="configure-the-manifest"></a>マニフェストを構成する

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xmlを開ソリューション エクスプローラー** 

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Visual Studio コードで、 **manifest.xml** ファイルを開きます。

---

1. マニフェストの下部付近には、終了 `</Resources>` 要素があります。 要素のすぐ下に、終了`</VersionOverrides>`要素の前に`</Resources>`次の XML を挿入します。 Outlook 以外の Office アプリケーションの場合は、セクションの末尾にマークアップを `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 追加します。 Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。

   ```xml
   <WebApplicationInfo>
       <Id>[application-id]</Id>
       <Resource>api://localhost:[port]/[application-id]</Resource>
       <Scopes>
           <Scope>openid</Scope>
           <Scope>user.read</Scope>
           <Scope>profile</Scope>
       </Scopes>
   </WebApplicationInfo>
   ```

1. プロジェクトの正しいポート番号に置き換えます `[port]` 。 office を使用してアドインを作成した場合、通常、ポート番号は 3000 であり、package.json ファイルにあります。 Visual Studio 2019 でアドインを作成した場合、ポートは Web プロジェクトの **SSL URL** プロパティにあります。
1. 両方 `[application-id]` のプレースホルダーを、アプリ登録の実際のアプリケーション ID に置き換えます。
1. ファイルを保存します。

挿入した XML には、次の要素と情報が含まれています。

- **\<WebApplicationInfo\>** - 次の要素の親。
- **\<Id\>** - アドインのクライアント ID これは、アドインの登録の一環として取得するアプリケーション ID です。 詳細については、「[Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。
- **\<Resource\>** - アドインの URL。 これは、AAD にアドインを登録したときに使用したのと同じ URI (`api:` プロトコルを含む) です。 この URI のドメイン部分は、アドインのマニフェストのセクションの URL で **\<Resources\>** 使用されるサブドメインを含むドメインと一致する必要があり、URI はクライアント ID で **\<Id\>** 終わる必要があります。
- **\<Scopes\>** - 1 つ以上 **\<Scope\>** の要素の親。
- **\<Scope\>** - アドインが AAD に必要なアクセス許可を指定します。 `profile` と `openID` のアクセス許可は常に必要です。ご利用のアドインが Microsoft Graph にアクセスしない場合、これは唯一必要なアクセス許可になる場合があります。 その場合は、必要な Microsoft Graph アクセス許可の要素も必要 **\<Scope\>** です。たとえば、`User.Read`. `Mail.Read` コードで使用している、Microsoft Graph にアクセスするためのライブラリでは、他にもアクセス許可が必要な場合があります。 たとえば、.NET 用の Microsoft 認証ライブラリ (MSAL) では、`offline_access` のアクセス許可が必要です。 詳細については、「[Office アドインで Microsoft Graph へ承認](authorize-to-microsoft-graph.md)」を参照してください。

## <a name="add-the-jwt-decode-package"></a>jwt デコード パッケージを追加する

API を `getAccessToken` 呼び出して、Office から ID トークンを取得できます。 まず、jwt デコード パッケージを追加して、ID トークンのデコードと表示を容易にします。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Visual Studio ソリューションを開きます。
1. メニューの [ツール] **> [NuGet パッケージ マネージャー] > [パッケージ マネージャー コンソール]** を選択します。
1. **パッケージ マネージャー コンソール** で次のコマンドを入力します。

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. ターミナル/コンソール ウィンドウから、アドイン プロジェクトのルート フォルダーに移動します。
1. 次のコマンドを入力します

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>作業ウィンドウに UI を追加する

ID トークンから取得するユーザー情報を表示できるように作業ウィンドウを変更する必要があります。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Home.html ファイルを開きます。
1. 次のスクリプト タグを `<head>` ページのセクションに追加します。 これには、前に追加した jwt デコード パッケージが含まれます。

   ```html
   <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
   ```

1. セクションを次の `<body>` HTML に置き換えます。

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. **src/taskpane/taskpane.html** ファイルを開きます。
1. セクションを次の `<body>` HTML に置き換えます。

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

---

## <a name="call-the-getaccesstoken-api"></a>getAccessToken API を呼び出す

最後の手順は、ID トークンを呼び出 `getAccessToken`して取得することです。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. **Home.js** ファイルを開きます。
1. ファイルのすべての内容を次のコードで置き換えます。

   ```javascript
   (function () {
     "use strict";

     // The initialize function must be run each time a new page is loaded.
     Office.initialize = function (reason) {
       $(document).ready(function () {
         $("#getIDToken").click(getIDToken);
       });
     };

     async function getIDToken() {
       try {
         let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
           allowSignInPrompt: true,
         });
         let userToken = jwt_decode(userTokenEncoded);
         document.getElementById("userInfo").innerHTML =
           "name: " +
           userToken.name +
           "<br>email: " +
           userToken.preferred_username +
           "<br>id: " +
           userToken.oid;
         console.log(userToken);
       } catch (error) {
         document.getElementById("userInfo").innerHTML =
           "An error occurred. <br>Name: " +
           error.name +
           "<br>Code: " +
           error.code +
           "<br>Message: " +
           error.message;
         console.log(error);
       }
     }
   })();
   ```

1. ファイルを保存します。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. **src/taskpane/taskpane.js** ファイルを開きます。
1. ファイルのすべての内容を次のコードで置き換えます。

   ```javascript
   import jwt_decode from "jwt-decode";

   Office.onReady((info) => {
     if (info.host === Office.HostType.Excel) {
       document.getElementById("getIDToken").onclick = getIDToken;
     }
   });

   async function getIDToken() {
     try {
       let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
         allowSignInPrompt: true,
       });
       let userToken = jwt_decode(userTokenEncoded);
       document.getElementById("userInfo").innerHTML =
         "name: " +
         userToken.name +
         "<br>email: " +
         userToken.preferred_username +
         "<br>id: " +
         userToken.oid;
       console.log(userToken);
     } catch (error) {
       document.getElementById("userInfo").innerHTML =
         "An error occurred. <br>Name: " +
         error.name +
         "<br>Code: " +
         error.code +
         "<br>Message: " +
         error.message;
       console.log(error);
     }
   }
   ```

1. ファイルを保存します。

---

## <a name="run-the-add-in"></a>アドインを実行する

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. [ **デバッグ] > [デバッグの開始]** を選択するか、 **F5** キーを押します。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

コマンド ラインから実行 `npm start` します。

---

1. Excel が起動したら、アプリ登録の作成に使用したのと同じテナント アカウントを使用して Office にサインインします。
1. **[ホーム**] リボンで 、[**タスクウィンドウの表示**] を選択してアドインを開きます。
1. アドインの作業ウィンドウで、[ **ID トークンの取得**] を選択します。

アドインには、サインインしたアカウントの名前、電子メール、ID が表示されます。

> [!NOTE]
> エラーが発生した場合は、この記事の登録手順でアプリの登録を確認してください。 アプリの登録を設定するときに詳細が見つからないことが、SSO での作業に関する一般的な原因です。 アドインを正常に実行できない場合は、 [シングル サインオン (SSO) のエラー メッセージのトラブルシューティングに関するページを](troubleshoot-sso-in-office-add-ins.md)参照してください。

## <a name="see-also"></a>関連項目

[要求を使用してユーザーを確実に識別する (サブジェクト ID とオブジェクト ID)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)

