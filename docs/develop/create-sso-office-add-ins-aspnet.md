---
title: シングル サインオンを使用する ASP.NET Office アドインを作成する
description: シングル サインオン (SSO) を使用するために、ASP.NET バックエンドを使用して Office アドインを作成 (または変換) する方法の詳細なガイド。
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b0179429f9d81b893394278580b6ef8891dd0a87
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/28/2022
ms.locfileid: "68842104"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する ASP.NET Office アドインを作成する

ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。 概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。
この記事では、ASP.NET で構築されたアドインでシングル サインオン (SSO) を有効にするプロセスについて説明します。

## <a name="prerequisites"></a>前提条件

- Visual Studio 2019 以降。

- Visual Studio を構成するときの **Office/SharePoint 開発** ワークロード。

- [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

- Microsoft 365 サブスクリプションのOneDrive for Businessに保存されている少なくともいくつかのファイルとフォルダー。

- アクティブなサブスクリプションを持つ Azure アカウント - [無料でアカウントを作成](https://azure.microsoft.com/free/?WT.mc_id=A261C142F)します。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

「[Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)」にあるリポジトリを複製するかダウンロードします。

> [!NOTE]
> サンプルには 2 つのバージョンがあります。
>
> - The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
> - このサンプルの **[Complete]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。 完成済みのバージョンを使用する場合は、この記事の手順をそのまま実行しますが、[Before] を [Complete] に置き換えて、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションを省略してください。

以降のアプリ登録手順のプレースホルダーには、次の値を使用します。

| プレースホルダー           | 値                                           |
|-----------------------|-------------------------------------------------|
| `<add-in-name>`       | **Office-Add-in-ASPNET-SSO**                    |
| `<redirect-platform>` | **Web**                                         |
| `<redirect-uri>`      | `https://localhost:44355/AzureADAuth/Authorize` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-solution"></a>ソリューションを構成する

1. [**Before**] フォルダーのルートで、**Visual Studio** でソリューション (.sln) ファイルを開きます。 [**ソリューション エクスプローラー**] の一番上のノード (プロジェクト ノードではなく、ソリューション ノード) を右クリックして、[**スタートアップ プロジェクトの設定**] を選択します。

1. [**共通プロパティ**] で、[**スタートアップ プロジェクト**]、[**マルチ スタートアップ プロジェクト**] の順に選択します。 両方のプロジェクトの [**アクション**] が [**開始**] に設定され、「... WebAPI」で終わるプロジェクトが最初にリストされていることを確認します。 ダイアログを閉じます。

1. **ソリューション エクスプローラー** に戻り、**Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトを選択 (右クリックしない) します。 [**プロパティ**] ウィンドウを開きます。 [**SSL 有効**] が [**True**] であることを確認します。 [**SSL URL**] が `http://localhost:44355/` であることを確認します。

1. 「Web.config」 で、以前にコピーした値を使用します。 [**ida:ClientID**] と [**ida:Audience**] の両方を [**アプリケーション (クライアント) ID**] に設定し、[**ida:Password**] をクライアント シークレットに設定します。 また、 **ida:Domain** を に `http://localhost:44355` 設定します (末尾にスラッシュ "/" はありません)。

    > [!NOTE]
    > **アプリケーション (クライアント) ID** は、Office クライアント アプリケーション (PowerPoint、Word、Excel など) などの他のアプリケーションがアプリケーションへの承認されたアクセスを求める場合の "対象ユーザー" の値です。 また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. アドインを登録したときに、**サポートされているアカウントの種類** で「この組織のディレクトリ内のアカウントのみ」を選択しなかった場合は、web.config を保存して閉じます。 それ以外の場合は、保存して、開いたままにします。

1. 引き続き **ソリューション エクスプローラー** で、**Office-Add-in-ASPNET-SSO** プロジェクトを選択し、アドイン マニフェスト ファイル "Office-Add-in-ASPNET-SSO.xml" を開き、ファイルの一番下までスクロールします。 終了 `</VersionOverrides>` タグのすぐ上に、次のマークアップがあります。

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. このマークアップ内の *両方の場所の* プレースホルダー「$application_GUID here$」を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。 「$」は ID の一部ではないので、これらを含めないでください。 これは、web.config の ClientID と Audience に使用したものと同じ ID です。

    > [!NOTE]
    > 値は **\<Resource\>** 、アドインを登録したときに設定した **アプリケーション ID URI** です。 セクションは **\<Scopes\>** 、アドインが AppSource を通じて販売されている場合にのみ、同意ダイアログ ボックスを生成するために使用されます。

1. ファイルを保存して閉じます。

### <a name="setup-for-single-tenant"></a>シングルテナントのセットアップ

アドインの登録時に [この組織のディレクトリ内のアカウントのみ] を **[サポートされているアカウントの種類** ] に選択した場合は、これらの追加のセットアップ手順を実行する必要があります。

1. Azure ポータルに戻り、アドインの登録の [**概要**] ブレードを開きます。 [**Directory (テナント) ID**] をコピーします。

1. web.config で、[**ida：Authority**] の値の「Common」を前の手順でコピーした GUID に置き換えます。 終了すると、値は `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />` のようになります。

1. web.config を保存して閉じます。

## <a name="code-the-client-side"></a>クライアント側のコードの作成

1. [**スクリプト**] フォルダー内の HomeES6.js ファイルを開きます。 既にいくつかのコードが含まれます。

    - Office が UI に Internet Explorer を使用しているときにアドインを実行できるように、Office.Promise オブジェクトをグローバル ウィンドウ オブジェクトに割り当てるポリフィル。 (詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。)
    - 次に `Office.initialize` 、ボタンクリック イベントにハンドラーを割り当てる関数への `getGraphAccessTokenButton` 割り当て。
    - `showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。
    - `logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。
    - SSO がサポートされていない、またはエラーが発生したシナリオでアドインが使用するフォールバック承認システムを実装するコード。

1. への割り当ての後に `Office.initialize`、次のコードを追加します。 このコードについては、以下の点に注意してください。

    - アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。 カウンター変数 `retryGetAccessToken` は、ユーザーがトークンを取得しようとしたときに繰り返し再試行されないように使用されます。
    - `getGraphData` 関数は、ES6 `async` キーワードで定義されます。 ES6 構文を使用すると、Office アドインの SSO API の使用が非常に簡単になります。 これは、ソリューション内の、Internet Explorer でサポートされていない構文を使用する唯一のファイルです。 ファイル名に「ES6」というリマインダーが設定されています。 このソリューションでは、tsc トランスパイラーを使用してこのファイルを ES5 にトランスパイルします。これにより、Office が UI に Internet Explorer を使用しているときにアドインが実行されます。 (プロジェクトのルートにある tsconfig.json ファイルを参照します。)

    ```javascript
    let retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. 関数の後に `getGraphData` 、次の関数を追加します。 後の手順で `handleClientSideErrors` 関数を作成することに注意してください。

    > [!NOTE]
    > この記事で使用する 2 つのアクセス トークンを区別するために、getAccessToken() から返されるトークンはブートストラップ トークンと呼ばれます。 その後、On-Behalf-Of フローを通じて、Microsoft Graph へのアクセス権を持つ新しいトークンと交換されます。

    ```javascript
    async function getDataWithToken(options) {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for a new access token to Microsoft Graph and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```


1. を次のコードに置き換えて `TODO 1` 、Office ホストからアクセス トークンを取得します。 *options* パラメーターには、前`getGraphData()`の関数から渡された次の設定が含まれています。

    - `allowSignInPrompt` が true に設定されています。 これにより、ユーザーがまだ Office にサインインしていない場合は、サインインするようにユーザーに求めるメッセージが Office に指示されます。
    - `allowConsentPrompt` が true に設定されています。 これにより、同意がまだ付与されていない場合は、アドインにユーザーのMicrosoft Azure Active Directory プロファイルへのアクセスを許可する同意を求めるメッセージが表示されます。 (結果のプロンプトでは、ユーザーが Microsoft Graph スコープに同意 *することはできません* )。
    - `forMSGraphAccess` が true に設定されています。 これにより、ユーザーまたは管理者がアドインの Graph スコープへの同意を付与していない場合に、エラー (コード 13012) が返されます。 Microsoft Graph にアクセスするには、アドインは、代わりにフローを介して新しいアクセス トークンのアクセス トークンを交換する必要があります。 を true に設定 `forMSGraphAccess` すると、 **getAccessToken()** が成功したが、Microsoft Graph の後で代理フローが失敗するシナリオを回避できます。 アドインのクライアント側コードが 13012 に返信するには、フォールバック認証システムに分岐します。

    また、次のコードにも注意してください。

    - 後の手順で `getData` 関数を作成します。
    - パラメーターは `/api/values` 、サーバー側コントローラーの URL で、代わりにフローを使用して、トークンを新しいアクセス トークンと交換して Microsoft Graph を呼び出します。

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken(options);

    getData("/api/values", bootstrapToken);
    ```

1. 関数の後に `getGraphData` 、次を追加します。 このコードについては、以下の点に注意してください。

    - これは、SSO 認証システムおよびフォールバック認証システムの両方で使用されます。
    - `relativeUrl` パラメーターは、サーバー側のコントローラーです。
    - `accessToken` パラメーターは、ブートストラップ トークンまたはフル アクセス トークンにすることができます。
    - `writeFileNamesToOfficeDocument` は、既にプロジェクトの一部です。
    - 後の手順で `handleServerSideErrors` 関数を作成します。

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a>クライアント側のエラーを処理する

1. 関数の後に `getData` 、次の関数を追加します。 `error.code`は数値であり、通常は 13xxx の範囲にあることを注意してください。

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. `TODO 2`を以下のコードに置き換えます。 これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-in's functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-in's functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. `TODO 3`を以下のコードに置き換えます。 その他のエラーが発生した場合、アドインはフォールバック認証システムに分岐します。 これらのエラーの詳細については、「 [Office アドインでの SSO のトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。このアドインでは、フォールバック システムによってダイアログが開き、ユーザーが既にサインインしている場合でもサインインする必要があります。

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>サーバー側のエラーを処理する

1. 関数の後に `handleClientSideErrors` 、次の関数を追加します。

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. `TODO 4` を以下のように置き換えます。 このコードについては、MFA などが存在する前に ASP.NET エラー クラスが作成されたことに注意してください。 第 2 認証要素に対する要求をサーバー側の論理が処理する方法の副作用として、クライアントに送信されるサーバー側のエラーは **Message** プロパティがありますが、**ExceptionMessage** プロパティはありません。 ただし、他のすべてのエラーには **ExceptionMessage** プロパティがあるため、クライアント側のコードは両方の応答を解析する必要があります。 どちらか一方の変数が未定義になります。

    ```javascript
    const message = JSON.parse(result.responseText).Message;
    const exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. `TODO 5` を以下のように置き換えます。 Microsoft Graph が認証の追加形式を必要とする場合、エラー AADSTS50076 が送信されます。 これには、**Message.Claims** プロパティの追加要件に関する情報が含まれます。 これを処理するために、コードはブートストラップ トークンの取得を 2 回試行しますが、今回は `authChallenge` オプションの値として追加要素の要求が含まれます。これにより、Azure AD は、必要なすべての形式の認証をユーザーに要求します。

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            const claims = JSON.parse(message).Claims;
            const claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. を次のように置き換えます `TODO 6` 。

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. `TODO 7` を以下のように置き換えます。 まれにブートストラップ トークンが Office の検証時に期限切れにならず、交換のために Azure AD に送信されるまでの間に期限切れになることがあることに注意してください。 Azure AD はエラー AADSTS500133 で応答します。 この場合、コードは SSO API を呼び戻します (ただし、1 回のみ)。 今回は、Office が期限切れになっていない新しいブートストラップ トークンを返します。

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. を次のように置き換えます `TODO 8` 。

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. ファイルを保存します。

## <a name="code-the-server-side"></a>サーバー側のコードを作成する

### <a name="configure-the-owin-middleware"></a>OWIN ミドルウェアを構成する

1. **Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトのルートにある Startup.cs ファイルを開き、**スタートアップ** クラスに次のメソッドを追加します。 `ConfigureAuth` メソッドは、この後の手順で作成することに注意してください。

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. ファイルを保存して閉じます。

1. **App_Start** フォルダーを右クリックして、**[追加] > [クラス]** を選択します。

1. **[新しい項目の追加]** ダイアログで、ファイルに「**Startup.Auth.cs**」という名前を付けて **[追加]** をクリックします。

1. 新しいファイルで名前空間の名前を `Office_Add_in_ASPNET_SSO_WebAPI` に短縮します。

1. ファイルの先頭に、次に示す `using` ステートメントがすべて揃っていることを確認します。

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:

    `public partial class Startup`

1. Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. `TODO 1` を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    - このコードは、Office アプリケーションから取得されるブートストラップ トークンで指定された対象ユーザーが、web.configで指定された値と一致する必要があることを OWIN に指示します。
    - Microsoft アカウントには、組織のテナント GUID とは異なる発行者 GUID があるため、両方の種類のアカウントをサポートするために、発行者は検証されません。
    - を に`true`設定`SaveSigninToken`すると、OWIN は Office アプリケーションから生のブートストラップ トークンを保存します。 これは、アドインが代理フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。
    - OWIN ミドルウェアでは、スコープは検証されません。 `access_as_user` が含まれている必要があるブートストラップ トークンのスコープは、コントローラーで検証されます。

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. `TODO 2`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    - より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。
    - メソッドに渡される URL は、OWIN ミドルウェアが、Office アプリケーションから受信したブートストラップ トークンの署名を確認するために必要なキーを取得するための手順を取得する場所です。 URL の権威セグメントは、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. ファイルを保存して閉じます。

### <a name="create-the-apivalues-controller"></a>/api/values コントローラーを作成する

1. ファイル **Controllers\ValueController.cs** を開きます。 このコントローラーは、SSO システムがブートストラップ トークンを正常に取得した場合に使用されます。 フォールバック認証システムの一部として使用されることはありません。 そのシステムで AzureADAuthController が使用されました。これは、自動的に作成されます。

1. ファイルの先頭に、次に示す `using` ステートメントがあることを確認します。

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.

1. 次のメソッドを `ValuesController` に追加します。 戻り値は、`Task<IEnumerable<string>>` ではなく `GET api/values` メソッドでより一般的な `Task<HttpResponseMessage>` になる点に注意してください。 これは、OAuth 承認ロジックが、ASP.NET フィルターではなくコントローラーに存在する必要があるという事実の副作用です。 その論理の一部のエラーの条件では、アドインのクライアントに HTTP 応答オブジェクトが送信される必要があります。

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //         token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get a new access token for Microsoft Graph.

        // TODO 4: Use the new access token to call Microsoft Graph.
    }
    ```

1. `TODO1` を次のコードに置き換えて、`access_as_user` を含むトークンで指定されているスコープを検証します。 `SendErrorToClient` メソッドの第 2 パラメーターは、**Exception** オブジェクトです。 この場合、コードは `null` を渡します。これは、**Exception** オブジェクトが含まれていることで、生成される HTTP 応答には **Message** プロパティが含められなくなるためです。

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. `TODO 2` を次のコードに置き換えて、「代理」フローを使用して Microsoft Graph のトークンを取得するために必要なすべての情報を編成します。 このコードについては、以下の点に注意してください。

    - アドインは、Office アプリケーションとユーザーがアクセスする必要があるリソース (または対象ユーザー) の役割を果たさなくなりました。 この時点で、それ自体が Microsoft Graph にアクセスする必要があるクライアントになります。 は MSAL の「クライアント コンテキスト」オブジェクトになります。
    - MSAL.NET 3.x.x からは、`bootstrapContext` は単なるブートストラップ トークンです。
    - 権威は、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。
    - コードが を要求 `profile`した場合、MSAL はエラーをスローします。これは、Office クライアント アプリケーションがアドインの Web アプリケーションにトークンを取得するときにのみ実際に使用されます。 そのため、`Files.Read.All` のみが明示的に要求されます。

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:Domain"])
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. `TODO 3` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。 それが見つからなかった場合にのみ、Azure AD V2 エンドポイントで代理フローを開始します。
    - `MsalServiceException` 以外の種類の例外は、意図的にキャッチしていないため、`500 Server Error` メッセージとしてクライアントに伝達されます。

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. `TODO 3a`を以下のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - Microsoft Graph リソースが多要素認証を必要としているときに、その認証をユーザーがまだ指定していない場合、Azure AD はエラー `AADSTS50076` と **Claims** プロパティを含む「400 要求が正しくありません」を返します。 MSAL は、この情報と共に **MsalUiRequiredException** (**MsalServiceException** から継承) をスローします。
    - **Claims** プロパティの値は、Office アプリケーションに渡すクライアントに渡す必要があります。この値は、新しいブートストラップ トークンの要求に含まれます。 Azure AD は、認証のすべての要求されたフォームをユーザーに示します。
    - The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.
    - カスタム メッセージは、JSON として書式設定されているため、クライアント側の JavaScript は既知の JavaScript `JSON` オブジェクトのメソッドでメッセージを解析できます。

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. `TODO 3b`を以下のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - Azure AD の呼び出しにユーザーまたはテナント管理者のどちらも同意していない (または同意が取り消された) スコープ (アクセス許可) が少なくとも 1 つ含まれていると、Azure AD はエラー `AADSTS65001` と共に「400 要求が正しくありません」を返します。 MSAL は、この情報と共に **MsalUiRequiredException** をスローします。
    - Azure AD の呼び出しに Azure AD が認識しないスコープが少なくとも 1 つ含まれていると、AAD はエラー `AADSTS70011` と共に「400 要求が正しくありません」を返します。 MSAL は、この情報と共に **MsalUiRequiredException** をスローします。
    - すべての説明が含まれている理由は、別の条件で 70011 が返されたときに、このアドインでは無効なスコープの存在を意味する場合のみを処理する必要があるためです。
    - The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. `TODO 3c` を次のコードに置き換えて、他のすべての **MsalServiceException** を処理します。

    ```csharp
    else
    {
        throw e;
    }
    ```

1. `TODO 4` を次のコードに置き換えます。 `GraphApiHelper.GetOneDriveFileNames` メソッドは、自動的に作成されます。これは、Microsoft Graph にデータを要求し、アクセス トークンを含めます。

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. ファイルを保存して閉じます。

## <a name="run-the-solution"></a>ソリューションを実行する

1. Visual Studio ソリューション ファイルを開きます。
1. [**ビルド**] メニューで [**ソリューションのクリーン**] を選択します。 終了したら、[**ビルド**] メニューをもう一度開き、[**ソリューションのビルド**] を選択します。
1. [**ソリューション エクスプローラー**] で、[**Office-Add-in-ASPNET-SSO**] を選択します (一番上のソリューション ノードではなく、「WebAPI」で終わる名前のプロジェクトではありません)。
1. [**プロパティ**] ウィンドウで、[**ドキュメントの開始**] ドロップダウンを開き、3 つのオプション (Excel、Word、または PowerPoint) のいずれかを選択します。

    ![目的の Office クライアント アプリケーション (Excel、PowerPoint、または Word) を選択します。](../images/SelectHost.JPG)

1. F5 キーを押します。
1. Office アプリケーションの [**ホーム**] リボンで、[**SSO ASP.NET**] グループの [**アドインの表示**] を選択して、タスク ウィンドウ アドインを開きます。
1. [**OneDrive ファイル名の取得**] ボタンをクリックします。 Microsoft 365 Educationまたは職場アカウント、または Microsoft アカウントを使用して Office にログインしていて、SSO が期待どおりに機能している場合は、作業ウィンドウにOneDrive for Businessの最初の 10 個のファイルとフォルダー名が表示されます。 ログインしていない場合、または SSO をサポートしていないシナリオ、または何らかの理由で SSO が機能しない場合は、サインインするように求められます。 サインインすると、ファイル名とフォルダー名が表示されます。

### <a name="testing-the-fallback-path"></a>フォールバック パスのテスト

フォールバック承認パスをテストするには、次の手順で SSO パスを強制的に失敗します。

1. 次のコードを、HomeES6.js ファイル内の メソッドの `getDataWithToken` 一番上に追加します。

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. 次に、 の呼び出しのすぐ上にある、同じメソッドのブロックの `try` 先頭に次の行を `getAccessToken`追加します。

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>ステージングと運用環境に移動するときのアドインの更新

すべての Office Web アドインと同様に、ステージング サーバーまたは運用サーバーに移行する準備ができたら、マニフェスト内のドメインを `localhost:44355` 新しいドメインで更新する必要があります。 同様に、web.config ファイル内のドメインを更新する必要があります。

ドメインは AAD 登録に表示されるため、新しいドメインが表示される場所の代わりに `localhost:44355` 使用するように登録を更新する必要があります。
