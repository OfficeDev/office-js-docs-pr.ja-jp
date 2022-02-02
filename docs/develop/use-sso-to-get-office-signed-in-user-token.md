---
title: SSO を使用してサインインしているユーザーの ID を取得する
description: getAccessToken API を呼び出して、サインインしているユーザーに関する名前、電子メール、追加情報を含む ID トークンを取得します。
ms.date: 01/25/2022
localization_priority: Normal
ms.openlocfilehash: 2c9b3c89a154d624f99e196014c7d8024286d927
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/02/2022
ms.locfileid: "62322337"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>SSO を使用してサインインしているユーザーの ID を取得する

API を`getAccessToken`使用して、ユーザーにサインインしている現在のユーザーの ID を含むアクセス トークンをOffice。 アクセス トークンは、サインインしているユーザーに関する ID クレーム (名前や電子メールなど) が含まれているため、ID トークンです。 ID トークンを使用して、独自の Web サービスを呼び出す際にユーザーを識別できます。 呼び出`getAccessToken`しを行う場合は、Officeで SSO を使用するアドインを構成するOffice。

この記事では、ID トークンを取得し、作業ウィンドウにユーザーの名前、電子メール、および一意の ID を表示する Office アドインを作成します。

> [!NOTE]
> SSO と Office API は`getAccessToken`、すべてのシナリオで機能しません。 SSO が使用できない場合は、常にフォールバック ダイアログを実装してユーザーにサインインします。 詳細については、「認証と承認[」を参照してください。Office API を使用します](auth-with-office-dialog-api.md)。

## <a name="create-an-app-registration"></a>アプリの登録を作成する

Office で SSO を使用するには、Microsoft ID プラットフォーム が Office アドインとそのユーザーに認証および承認サービスを提供できるよう、Azure portal でアプリ登録を作成する必要があります。

1. アプリを登録するには、 [Azure portal - アプリ登録ページに移動](https://go.microsoft.com/fwlink/?linkid=2083908) します。

1. 管理者資格情報を **_使用してテナント_** にサインインMicrosoft 365します。 たとえば、MyName@contoso.onmicrosoft.com です。

1. **[新規登録]** を選択します。 **[アプリケーションを登録]** ページで、次のように値を設定します。

   - `Office-Add-in-SSO` に **[名前]** を設定します。
   - **[サポートされているアカウントの種類]** を **[任意の組織のディレクトリ内のアカウントと個人用の Microsoft アカウント (例: Skype、 Xbox、Outlook.com)]** に設定します。
   - アプリケーションの種類を Web に設定 **し** 、[リダイレクト **URI] をに設定** します `https://localhost:[port]/dialog.html`。 Web `[port]` アプリケーションの正しいポート番号に置き換える。 yo office を使用してアドインを作成した場合、ポート番号は通常 3000 で、package.json ファイルに含まれています。 Visual Studio 2019 年に 2019 年にアドインを作成した場合、ポートは Web プロジェクトの **SSL URL** プロパティにあります。
   - **[登録]** を選択します。

1. [Office **-アドイン SSO**] ページで、アプリケーション **(クライアント) ID** とディレクトリ **(テナント) ID** の値をコピーして保存します。 以降の手順では、それらの両方を使用します。

   > [!NOTE]
   > この **アプリケーション (クライアント) ID** は、Office クライアント アプリケーション (PowerPoint、Word、Excel など) などの他のアプリケーションがアプリケーションへの承認されたアクセスを求める場合の "対象ユーザー" 値です。 また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. **[管理]** の下の **[認証]** を選択します。 [暗黙的な **付与] セクション** で、Access トークンと ID トークン **の両方のチェック ボックスを****有効にします**。

1. フォームの最上部で **[保存]** を選択します。

1. **[管理]** の下の **[API の公開]** を選択します。 [設定] **リンクを選択** します。 これにより、アプリケーション (クライアント) `api://[app-id-guid]``[app-id-guid]` ID という形式の **アプリケーション ID URI が生成されます**。

1. 生成された ID で、ダブル `localhost:[port]/` スラッシュと GUID の間に挿入 (末尾にスラッシュ "/" が付加されている点に注意してください)。 Web `[port]` アプリケーションの正しいポート番号に置き換える。 yo office を使用してアドインを作成した場合、ポート番号は通常 3000 で、package.json ファイルに含まれています。 Visual Studio 2019 年に 2019 年にアドインを作成した場合、ポートは Web プロジェクトの **SSL URL** プロパティにあります。
   完了したら、ID `api://localhost:[port]/[app-id-guid]`全体にフォーム (たとえば) が必要です `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

1. **[Scope の追加]** ボタンをクリックします。 開いたパネルで、`access_as_user`を **[スコープ名]** として入力します。

1. **[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。

1. `access_as_user`管理者とユーザーの同意のプロンプトを構成するためのフィールドに、Office クライアント アプリケーションが現在のユーザーと同じ権限でアドインの Web API を使用できる範囲に適した値を入力します。 提案:

   - **管理者の同意表示名**: Officeユーザーとして機能できます。
   - **管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。
   - **ユーザーの同意表示名**: Officeとして機能する場合があります。
   - **ユーザーの同意の** 説明: Officeと同じ権限を持つアドインの Web API を呼び出す方法を有効にしてください。

1. **[状態]** が **[有効]** に設定されていることを確認してください。

1. **[スコープの追加]** を選択します。

   > [!NOTE]
   > テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。

1. **[承認済みのクライアント アプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。 次のそれぞれの ID を事前承認する必要があります。

   - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
   - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
   - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)
   - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

   ID ごとに、次の手順を実行します。

   a. [ **クライアント アプリケーションの追加]** ボタン `[app-id-guid]` を選択し、開くパネルで、アプリケーション (クライアント) ID に設定し、 のチェック ボックスをオンにします `api://localhost:44355/[app-id-guid]/access_as_user`。

   b. **[アプリケーションの追加]** を選択します。

1. **[管理]** の下の **[API アクセス許可]** を選択し、**[アクセス許可の追加]** を選択します。 開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。

1. アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。 プロファイルのアクセス許可を検索して **選択** します。 この`profile`アクセス許可は、Office Web アプリケーションにトークンを取得するために必要です。

   - profile

   > [!NOTE]
   > `User.Read` アクセス許可は既定でリストされています。 必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。

1. パネル下部にある **[アクセス許可の追加]** ボタンを選択します。

1. 同じページで、[ **管理者の同意を \<tenant-name\>** 許可する] ボタンを選択し、表示される確認のために **[は** い] を選択します。

## <a name="create-the-office-add-in"></a>Office アドインを作成する

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 2019 Visual Studioを開始し、[新しいプロジェクトの作成 **] を選択します**。
1. Web アドイン プロジェクト テンプレートExcel **を検索して** 選択します。 **[次へ]** を選択します。 注: SSO は任意のOfficeアプリケーションで動作しますが、この記事では、このアプリケーションでExcel。
1. **sso-display-user-info** などのプロジェクト名を入力し、[作成] を選択 **します**。 他のフィールドは既定値のままにできます。
1. [アドイン **の種類の選択]** ダイアログ ボックスで、[新しい機能を追加する] を選択し、[**Excel] を** 選択 **します**。

プロジェクトが作成され、ソリューションに 2 つのプロジェクトが含まれています。

- **sso-display-user-info**: アドインをサイドローディングするマニフェストと詳細がExcel。
- **sso-display-user-infoWeb**: アドイン ASP.NET Web ページをホストするプロジェクトの一覧です。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

必ず [開発 [環境をセットアップする] を選択します](../overview/set-up-your-dev-environment.md)。

1. 次のコマンドを入力してプロジェクトを作成します。

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

プロジェクトは、 **sso-display-user-info という名前の新しいフォルダーに作成されます**。

---

## <a name="configure-the-manifest"></a>マニフェストを構成する

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. ソリューション **エクスプローラーで** **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. コードVisual Studio、ファイルを **開** manifest.xmlします。

---

1. マニフェストの下部の近くに閉じる要素 `</Resources>` があります。 要素の下に、終了要素の前に `</Resources>` 次の XML を挿入 `</VersionOverrides>` します。 [Office以外のアプリケーションOutlook、セクションの末尾にマークアップを追加`<VersionOverrides ... xsi:type="VersionOverridesV1_0">`します。 Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。

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

1. プロジェクト `[port]` の正しいポート番号に置き換える。 yo office を使用してアドインを作成した場合、ポート番号は通常 3000 で、package.json ファイルに含まれています。 Visual Studio 2019 年に 2019 年にアドインを作成した場合、ポートは Web プロジェクトの **SSL URL** プロパティにあります。
1. 両方のプレースホルダー `[application-id]` を、アプリ登録の実際のアプリケーション ID に置き換える。
1. ファイルを保存します。

挿入した XML には、次の要素と情報が含まれています。

- **WebApplicationInfo** - 次の要素の親。
- **Id** -このアドインのクライアント ID。これはアドインを登録する一貫として取得するアプリケーション ID です。 詳細については、「[Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。
- **Resource** - アドインの URL。 これは、AAD にアドインを登録したときに使用したのと同じ URI (`api:` プロトコルを含む) です。 この URI のドメイン部分は、アドインのマニフェストの `<Resources>` のセクションの URL で使用されている任意のサブドメインを含むドメインと一致し、URI の末尾が `<Id>` 内のクライアント ID で終了している必要があります。
- **Scopes** - 1 つ以上の **Scope** 要素の親。
- **Scope** - アドインが AAD に対して必要なアクセス許可を指定する。 `profile` と `openID` のアクセス許可は常に必要です。ご利用のアドインが Microsoft Graph にアクセスしない場合、これは唯一必要なアクセス許可になる場合があります。 アクセスする場合、Microsoft Graph へのアクセスに必要な許可として、`User.Read`、`Mail.Read` など **Scope** 要素も必要になります。 コードで使用している、Microsoft Graph にアクセスするためのライブラリでは、他にもアクセス許可が必要な場合があります。 たとえば、.NET 用の Microsoft 認証ライブラリ (MSAL) では、`offline_access` のアクセス許可が必要です。 詳細については、「[Office アドインで Microsoft Graph へ承認](authorize-to-microsoft-graph.md)」を参照してください。

## <a name="add-the-jwt-decode-package"></a>jwt-decode パッケージを追加する

API を呼び出して`getAccessToken`、ID トークンをユーザーから取得Office。 まず、jwt-decode パッケージを追加して、ID トークンのデコードと表示を容易にします。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. ソリューションをVisual Studioします。
1. メニューの [ツール] **メニューの [コンソール> NuGet パッケージ マネージャー > パッケージ マネージャーします**。
1. コンソールで次のコマンド **をパッケージ マネージャーします**。

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. ターミナル/コンソール ウィンドウから、アドイン プロジェクトのルート フォルダーに移動します。
1. 次のコマンドを入力します。

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>作業ウィンドウに UI を追加する

作業ウィンドウを変更して、ID トークンから取得するユーザー情報を表示する必要があります。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. ファイルを開Home.htmlします。
1. ページのセクションに次のスクリプト タグ `<head>` を追加します。 これには、前に追加した jwt-decode パッケージが含まれます。

   ```html
   <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
   ```

1. セクションを次 `<body>` の HTML に置き換える。

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

1. **src/taskpane/taskpane.htmlファイルを開** きます。
1. セクションを次 `<body>` の HTML に置き換える。

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

最後の手順は、呼び出して ID トークンを取得します `getAccessToken`。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. ファイルを **開Home.js** します。
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

1. **src/taskpane/taskpane.jsファイルを開** きます。
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

1. [ **デバッグ] を>を開始するか**、 **F5 キーを押します**。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

コマンド `npm start` ラインから実行します。

---

1. アプリExcel、アプリの登録Officeと同じテナント アカウントでサインインします。
1. [ホーム **] リボンで** 、[ **タスクウィンドウの表示] を選択して** アドインを開きます。
1. アドインの作業ウィンドウで、[GET ID トークン] **を選択します**。

アドインには、サインインしたアカウントの名前、電子メール、ID が表示されます。

> [!NOTE]
> エラーが発生した場合は、この記事のアプリ登録の登録手順を確認してください。 アプリの登録を設定するときに詳細が見つからないのは、SSO を操作する際の問題の一般的な原因です。 それでもアドインを正常に実行できない場合は、「シングル サインオン (SSO)のエラー メッセージのトラブルシューティング」を [参照してください](troubleshoot-sso-in-office-add-ins.md)。

## <a name="see-also"></a>関連項目

[クレームを使用してユーザーを確実に識別する (Subject および Object ID)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)
