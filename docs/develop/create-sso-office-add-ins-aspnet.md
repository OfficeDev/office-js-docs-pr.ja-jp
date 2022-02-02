---
title: シングル サインオンを使用する ASP.NET Office アドインを作成する
description: シングル サインオン (SSO) を使用する ASP.NET バックエンドを使用して Office アドインを作成 (または変換) する方法の詳細なガイド。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: a8d2cd20e9ad47e18ff6ee84cbd45c27f89d537c
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320257"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する ASP.NET Office アドインを作成する

ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。 概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。
この記事では、シングル サインオン (SSO) を使用してビルドされたアドインでシングル サインオン (SSO) を有効にするプロセス ASP.NET。

## <a name="prerequisites"></a>前提条件

* Visual Studio 2019 以降。

* Visual studio **をOffice場合SharePoint開発** ワークロードを管理する必要があります。

* [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* サブスクリプションに保存されているファイルとフォルダー OneDrive for Business少なくともMicrosoft 365です。

* アクティブなサブスクリプションを持つ Azure アカウント - [無料でアカウントを作成します](https://azure.microsoft.com/free/?WT.mc_id=A261C142F)。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

「[Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)」にあるリポジトリを複製するかダウンロードします。

> [!NOTE]
> サンプルには 2 つのバージョンがあります。
>
> * **[Before]** フォルダーはスタート プロジェクトです。SSO や承認に直接関連しない UI などの側面は、既に完了しています。この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。
> * このサンプルの **[Complete]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。 完成済みのバージョンを使用する場合は、この記事の手順をそのまま実行しますが、[Before] を [Complete] に置き換えて、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションを省略してください。

## <a name="register-the-add-in-through-an-app-registration"></a>アプリ登録を通じてアドインを登録する

最初に、「[クイック スタート:](/azure/active-directory/develop/quickstart-register-app) アプリケーションをアプリケーションに登録する」の手順をMicrosoft ID プラットフォームアドインを登録します。

アプリの登録には、次の設定を使用します。

* 名前: `Office-Add-in-ASPNET-SSO`
* サポートされているアカウントの種類: 任意の組織ディレクトリ内のアカウント **(Azure AD ディレクトリ - マルチテナント) と個人用 Microsoft アカウント (Skype、Xbox など)**

    > [!NOTE]
    >  アドインを登録するテナント内のユーザーだけがアドインを使用できる場合は、この組織ディレクトリの [アカウント] のみを選択できます **。** 代わりに、追加のセットアップ手順を実行する必要があります。 この記事 **の後半の「シングル テナントのセットアップ** 」を参照してください。

* プラットフォーム: **Web**
* リダイレクト URI: **https://localhost:44355/AzureADAuth/Authorize**
* クライアント シークレット: `*********` (作成後にこの値を記録する - 1 回だけ表示されます)

### <a name="expose-a-web-api"></a>Web API を公開する

1. 作成したアプリ登録で、[API を公開 **する] を選択し、[>を追加する] を選択します**。
   まだ構成していない場合は、 **アプリケーション ID URI** を設定するように求めるメッセージが表示されます。

    アプリ ID URI は、API のコードで参照するスコープのプレフィックスとして機能し、グローバルに一意である必要があります。 フォームを使用します`api://localhost:44355/[application-id-guid]`。`api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`

1. スコープの属性を [スコープの追加 **] ウィンドウで指定** します。

    |フィールド          |値  |
    |---------------|---------|
    |**スコープ名** | `access_as_user`|
    |**Who同意できる** | **管理者とユーザー**|
    |**管理者の同意表示名** | Officeユーザーとして機能できます。|
    |**管理者の同意の説明** | 現在Officeと同じ権限を持つアドインの Web API を呼び出す方法を有効にしてください。|
    |**ユーザーの同意表示名** | Officeは、自分として機能できます。|
    |**ユーザーの同意の説明** | ユーザー Office同じ権限を持つアドインの Web API を呼び出す方法を有効にしてください。|

1. [状態] **を [****有効] に設定** し、[スコープの追加] **を選択します**。

    > [!NOTE]
    > テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。

1. **[承認済みのクライアント アプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。 次のそれぞれの ID を事前承認する必要があります。

    |クライアント ID                              |アプリケーション  |
    |---------------------------------------|-----------------|
    |`d3590ed6-52b3-4102-aeff-aad2292ab01c` |Microsoft Office |
    |`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` |Microsoft Office |
    |`93d53678-613d-4013-afc1-62e9e444a0a5` |Office on the web |
    |`57fb890c-0dab-4253-a5e0-7188c88b2bb4` |Office on the web |
    |`08e18876-6177-487e-b8b5-cf950c1e598c` |Office on the web |
    |`bc59ab01-8403-45c6-8796-ac3ef710b3e3` |Outlook on the web |

    > [!NOTE]
    > ID ea5a67f6-b6f3-4338-b240-c655ddc3cc8e には、リストされている他のすべての ID が含まれています。Office アドイン SSO フローでサービスで使用する Office ホスト エンドポイントのすべてを事前承認するために、単独で使用できます。

    クライアント ID ごとに、次の手順を実行します。

    a. **[クライアント アプリケーションの追加]** を選択します。 開くパネルで、クライアント ID をそれぞれの GUID に設定し、 のチェック ボックスをオンにします `api://localhost:44355/[application-id-guid]/access_as_user`。

    b. **[アプリケーションの追加]** を選択します。

### <a name="configure-microsoft-graph-permissions"></a>Microsoft のアクセス許可Graph構成する

1. [**API アクセス許可] > Microsoft のアクセス許可を追加>を選択** Graph。

1. [**委任されたアクセス許可**] を選択します。 Microsoft Graphは多数のアクセス許可を公開します。最も一般的に使用されるアクセス許可は一覧の上部に表示されます。

1. [ **アクセス許可の選択] で**、次のアクセス許可を選択します。

    |アクセス許可     |説明  |
    |---------------|-------------|
    |Files.Read.All |ユーザーがアクセスできるすべてのファイルを読み取る。 |
    |profile        |ユーザーの基本的なプロファイルを表示します。 アドイン Web アプリケーションOfficeトークンを取得するには、アプリケーションが必要です。 |

    > [!NOTE]
    > `User.Read` アクセス許可は既定でリストされています。 必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。

1. [アクセス **許可の追加] を** 選択して、プロセスを完了します。

アクセス許可を構成するたびに、アプリのユーザーがサインイン時に、アプリが自分の代わりにリソース API にアクセスするための同意を求めるメッセージが表示されます。 管理者は、すべてのユーザーに代わって同意を与え、ユーザーが同意を求めないので、同意を許可できます。

1. 同じページで、[**[テナント名] に管理者の同意を与える**] ボタンを選択し、表示される確認に対して [**同意する**] を選択します。

    > [!NOTE]
    > [**[テナント名] に管理者の同意を与える**] を選択すると、同意プロンプトを作成できるように、数分後に再試行を求めるバナー メッセージが表示される場合があります。 その場合は、次のセクションで作業を開始できますが、ポータルに戻ってこのボタン **_を押してください_**。

## <a name="configure-the-solution"></a>ソリューションを構成する

1. [**Before**] フォルダーのルートで、**Visual Studio** でソリューション (.sln) ファイルを開きます。 [**ソリューション エクスプローラー**] の一番上のノード (プロジェクト ノードではなく、ソリューション ノード) を右クリックして、[**スタートアップ プロジェクトの設定**] を選択します。

1. [**共通プロパティ**] で、[**スタートアップ プロジェクト**]、[**マルチ スタートアップ プロジェクト**] の順に選択します。 両方のプロジェクトの [**アクション**] が [**開始**] に設定され、「... WebAPI」で終わるプロジェクトが最初にリストされていることを確認します。 ダイアログを閉じます。

1. ソリューション エクスプローラー **に戻り**、アドイン **ASPNET-SSO-WebAPI** プロジェクトOffice (右クリックしない) を選択します。 [**プロパティ**] ウィンドウを開きます。 [**SSL 有効**] が [**True**] であることを確認します。 [**SSL URL**] が `http://localhost:44355/` であることを確認します。

1. 「Web.config」 で、以前にコピーした値を使用します。 [**ida:ClientID**] と [**ida:Audience**] の両方を [**アプリケーション (クライアント) ID**] に設定し、[**ida:Password**] をクライアント シークレットに設定します。 また、 **ida:Domain を** `http://localhost:44355` (末尾にスラッシュ "/" なし) に設定します。

    > [!NOTE]
    > アプリケーション **(クライアント) ID** は、Office クライアント アプリケーション (PowerPoint、Word、Excel など) などの他のアプリケーションがアプリケーションへの承認アクセスを求める場合の "対象ユーザー" 値です。 また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. アドインを登録したときに、**サポートされているアカウントの種類** で「この組織のディレクトリ内のアカウントのみ」を選択しなかった場合は、web.config を保存して閉じます。 それ以外の場合は、保存して、開いたままにします。

1. ソリューション **エクスプローラーで**、**Office-Add-in-ASPNET-SSO** プロジェクトを選択し、アドイン マニフェスト ファイル "Office-Add-in-ASPNET-SSO.xml" を開き、ファイルの下部までスクロールします。 終了タグの上 `</VersionOverrides>` に、次のマークアップがあります。

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
    > **リソース** 値は、アドインを登録したときに設定した **アプリケーション ID URI** です。 **[範囲]** セクションは、アドインが AppSource を通じて販売される場合に同意ダイアログ ボックスを生成するためにのみ使用されます。

1. ファイルを保存して閉じます。

### <a name="setup-for-single-tenant"></a>シングルテナントのセットアップ

アドインの登録時にサポートされるアカウントの種類に対して [この組織ディレクトリ内のアカウントのみ] を選択した場合は、次の追加セットアップ手順を実行する必要があります。

1. Azure ポータルに戻り、アドインの登録の [**概要**] ブレードを開きます。 [**Directory (テナント) ID**] をコピーします。

1. web.config で、[**ida：Authority**] の値の「Common」を前の手順でコピーした GUID に置き換えます。 終了すると、値は `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />` のようになります。

1. web.config を保存して閉じます。

## <a name="code-the-client-side"></a>クライアント側のコードの作成

1. [**スクリプト**] フォルダー内の HomeES6.js ファイルを開きます。 既にいくつかのコードが含されています。

    * Office が UI に Internet Explorer を使用しているときにアドインを実行できるように、Office.Promise オブジェクトをグローバル ウィンドウ オブジェクトに割り当てるポリフィル。 (詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。)
    * `Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。
    * `showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。
    * `logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。
    * SSO がサポートされていないシナリオやエラーが発生したシナリオでアドインが使用するフォールバック承認システムを実装するコード。

1. `Office.initialize` への割り当ての下に、次に示すコードを追加します。 このコードについては、以下の点に注意してください。


    * アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。 カウンター変数 `retryGetAccessToken` は、ユーザーがトークンを取得しようとしたときに繰り返し再試行されないように使用されます。
    * `getGraphData` 関数は、ES6 `async` キーワードで定義されます。 ES6 構文を使用すると、Office アドインの SSO API の使用が非常に簡単になります。 これは、ソリューション内の、Internet Explorer でサポートされていない構文を使用する唯一のファイルです。 ファイル名に「ES6」というリマインダーが設定されています。 このソリューションでは、tsc トランスパイラーを使用してこのファイルを ES5 にトランスパイルします。これにより、Office が UI に Internet Explorer を使用しているときにアドインが実行されます。 (プロジェクトのルートにある tsconfig.json ファイルを参照します。)

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. `getGraphData` 関数の下に、次の関数を追加します。 後の手順で `handleClientSideErrors` 関数を作成することに注意してください。

    > [!NOTE]
    > この記事で使用する 2 つのアクセス トークンを区別するために、getAccessToken() から返されるトークンはブートストラップ トークンと呼ばれます。 後で、On-Behalf-Of フローを通じて、Microsoft サービスへのアクセス権を持つ新しいトークンGraph。

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


1. 次`TODO 1`のコードに置き換え、ホストからアクセス トークンOfficeします。 **options パラメーターには**、前の **getGraphData()** 関数から渡された次の設定が含まれます。

    * `allowSignInPrompt` は true に設定されます。 これにより、Officeがサインインしていない場合にサインインするように求めるメッセージが表示Office。
    * `allowConsentPrompt` は true に設定されます。 これにより、Officeがまだ許可されていない場合、アドインがユーザーの Microsoft Azure Active Directory プロファイルにアクセスすることへの同意を求めるメッセージが表示されます。 (結果のプロンプトでは *、* ユーザーが Microsoft のスコープに同意Graphしません)。
    * `forMSGraphAccess` は true に設定されます。 これにより、Officeまたは管理者がアドインの Graph スコープへの同意を許可していない場合、エラー (コード 13012) が返されます。 Microsoft にアクセスGraphアドインは、代理フローを介してアクセス トークンを新しいアクセス トークンと交換する必要があります。 true `forMSGraphAccess` に設定すると、**getAccessToken()** が成功したが、後で Microsoft サーバーの場合、代理フローが失敗するシナリオを回避Graph。 アドインのクライアント側コードが 13012 に返信するには、フォールバック認証システムに分岐します。

    また、次のコードに注意してください。

    * 後の手順で `getData` 関数を作成します。
    * この`/api/values`パラメーターは、サーバー側コントローラーの URL で、On-behalf-of flow を使用して、Microsoft Graph を呼び出す新しいアクセス トークンのトークンを交換します。

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken(options);

    getData("/api/values", bootstrapToken);
    ```

1. `getGraphData` 関数の下に、次を追加します。 このコードについては、以下の点に注意してください。

    * これは、SSO 認証システムおよびフォールバック認証システムの両方で使用されます。
    * `relativeUrl` パラメーターは、サーバー側のコントローラーです。
    * `accessToken` パラメーターは、ブートストラップ トークンまたはフル アクセス トークンにすることができます。
    * `writeFileNamesToOfficeDocument` は、既にプロジェクトの一部です。
    * 後の手順で `handleServerSideErrors` 関数を作成します。

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

1. `getData` 関数の下に、次の関数を追加します。 `error.code`は数値であり、通常は 13xxx の範囲にあることを注意してください。

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
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
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

1. `TODO 3`を以下のコードに置き換えます。 その他のエラーが発生した場合、アドインはフォールバック認証システムに分岐します。 これらのエラーの詳細については、「[トラブルシューティング SSO in Officeアドイン」を参照してください](troubleshoot-sso-in-office-add-ins.md)。このアドインでは、フォールバック システムによってダイアログが開き、ユーザーが既にサインインしている場合でも、ユーザーがサインインする必要があります。

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>サーバー側のエラーを処理する

1. `handleClientSideErrors` 関数の下に、次の関数を追加します。

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
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. `TODO 5` を以下のように置き換えます。 Microsoft Graph が認証の追加形式を必要とする場合、エラー AADSTS50076 が送信されます。 これには、**Message.Claims** プロパティの追加要件に関する情報が含まれます。 これを処理するために、コードはブートストラップ トークンの取得を 2 回試行しますが、今回は `authChallenge` オプションの値として追加要素の要求が含まれます。これにより、Azure AD は、必要なすべての形式の認証をユーザーに要求します。

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. `TODO 6` を以下のように置き換えます。

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

1. `TODO 8` を以下のように置き換えます。

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

1. `Startup` クラスの宣言にキーワード `partial` を追加します (まだ追加されていない場合)。これは、次のようになります。

    `public partial class Startup`

1. 次に示すメソッドを `Startup` クラスに追加します。このメソッドでは、クライアント側の Home.js ファイルの `getData` メソッドから渡されたアクセス トークンを OWIN ミドルウェアで検証する方法を指定します。承認プロセスは、`[Authorize]` 属性で修飾された Web API エンドポイントが呼び出されたときには必ずトリガーされます。

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. `TODO 1` を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * このコードは、OWIN に対して、Office アプリケーションから取得されるブートストラップ トークンで指定された対象ユーザーが、web.config で指定された値と一致するように指示します。
    * Microsoft アカウントには、組織のテナント GUID とは異なる発行者 GUID が含むので、両方の種類のアカウントをサポートするために、発行者を検証することはできません。
    * OWIN `SaveSigninToken` が`true`アプリケーションから未加工のブートストラップ トークンを保存Officeします。 これは、アドインが代理フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。
    * OWIN ミドルウェアでは、スコープは検証されません。 `access_as_user` が含まれている必要があるブートストラップ トークンのスコープは、コントローラーで検証されます。

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. `TODO 2`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。
    * メソッドに渡される URL は、OWIN ミドルウェアが、Office アプリケーションから受信したブートストラップ トークンの署名を確認するために必要なキーを取得する手順を取得します。 URL の権威セグメントは、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。

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

1. `ValuesController` を宣言している行のすぐ上に、属性 `[Authorize]` を追加します。これにより、アドインはコントローラー メソッドが呼び出されたときに、最後の手順で構成した承認プロセスを必ず実行するようになります。アドインへの有効なアクセス トークンを持つ呼び出し元のみが、コントローラーのメソッドを起動できます。

1. 次のメソッドを `ValuesController` に追加します。 戻り値は、`Task<IEnumerable<string>>` ではなく `GET api/values` メソッドでより一般的な `Task<HttpResponseMessage>` になる点に注意してください。 これは、OAuth 承認ロジックがコントローラー内にある必要があるという事実の副作用 ASP.NET です。 その論理の一部のエラーの条件では、アドインのクライアントに HTTP 応答オブジェクトが送信される必要があります。

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

    * アドインが、アプリケーションとユーザーがアクセスする必要があるリソース (または対象ユーザー) の役割Office果たしていなくなりました。 この時点で、それ自体が Microsoft Graph にアクセスする必要があるクライアントになります。 は MSAL の「クライアント コンテキスト」オブジェクトになります。
    * MSAL.NET 3.x.x からは、`bootstrapContext` は単なるブートストラップ トークンです。
    * 権威は、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。
    * MSAL は、`profile`Office クライアント アプリケーションがアドインの Web アプリケーションにトークンを取得するときにのみ使用される、コード要求の場合にエラーをスローします。 そのため、`Files.Read.All` のみが明示的に要求されます。

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

1. `TODO 3` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。 それが見つからなかった場合にのみ、Azure AD V2 エンドポイントで代理フローを開始します。
    * `MsalServiceException` 以外の種類の例外は、意図的にキャッチしていないため、`500 Server Error` メッセージとしてクライアントに伝達されます。

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

1. `TODO 3a` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * Microsoft Graph リソースが多要素認証を必要としているときに、その認証をユーザーがまだ指定していない場合、Azure AD はエラー `AADSTS50076` と **Claims** プロパティを含む「400 要求が正しくありません」を返します。 MSAL は、この情報と共に **MsalUiRequiredException** (**MsalServiceException** から継承) をスローします。
    * **Claims プロパティ** の値は、Office アプリケーションに渡す必要があるクライアントに渡す必要があります。この値は、新しいブートストラップ トークンの要求に含まれます。 Azure AD は、認証のすべての要求されたフォームをユーザーに示します。
    * 例外から HTTP 応答を作成する API は、**Claims** プロパティを認識しないため、このプロパティを応答オブジェクトに含めません。 これが含まれたメッセージを手動で作成する必要があります。 ただし、カスタムの **Message** プロパティは **ExceptionMessage** プロパティの作成を妨げるため、クライアントがエラー ID `AADSTS50076` を取得するには、その ID をカスタムの **Message** に追加する以外に方法はありません。 クライアントの JavaScript では、応答に **Message** または **ExceptionMessage** が含まれているかどうかを検出する必要があるため、どちらを読み取るかを認識します。
    * カスタム メッセージは、JSON として書式設定されているため、クライアント側の JavaScript は既知の JavaScript `JSON` オブジェクトのメソッドでメッセージを解析できます。

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. `TODO 3b` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * Azure AD の呼び出しにユーザーまたはテナント管理者のどちらも同意していない (または同意が取り消された) スコープ (アクセス許可) が少なくとも 1 つ含まれていると、Azure AD はエラー `AADSTS65001` と共に「400 要求が正しくありません」を返します。 MSAL は、この情報と共に **MsalUiRequiredException** をスローします。
    * Azure AD の呼び出しに Azure AD が認識しないスコープが少なくとも 1 つ含まれていると、AAD はエラー `AADSTS70011` と共に「400 要求が正しくありません」を返します。 MSAL は、この情報と共に **MsalUiRequiredException** をスローします。
    * すべての説明が含まれている理由は、別の条件で 70011 が返されたときに、このアドインでは無効なスコープの存在を意味する場合のみを処理する必要があるためです。
    * **MsalUiRequiredException** オブジェクトが `SendErrorToClient` に渡されます。これにより、エラー情報を格納している **ExceptionMessage** プロパティが HTTP 応答に含まれるようにします。

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

    ![目的のクライアント アプリケーションOffice選択します。Excel、PowerPoint、または Word。](../images/SelectHost.JPG)

1. F5 キーを押します。
1. Office アプリケーションの [**ホーム**] リボンで、[**SSO ASP.NET**] グループの [**アドインの表示**] を選択して、タスク ウィンドウ アドインを開きます。
1. [**OneDrive ファイル名の取得**] ボタンをクリックします。 Microsoft 365 Education または仕事用アカウント、または Microsoft アカウントを使用して Office にログインし、SSO が期待通り動作している場合、OneDrive for Business の最初の 10 ファイル名とフォルダー名が作業ウィンドウに表示されます。 ログインしていない場合、または SSO をサポートしていないシナリオの場合、または SSO が何らかの理由で動作しない場合は、サインインするように求めるメッセージが表示されます。 サインインすると、ファイル名とフォルダー名が表示されます。

### <a name="testing-the-fallback-path"></a>フォールバック パスのテスト

フォールバック承認パスをテストするには、次の手順で SSO パスを強制的に失敗します。

1. 次のコードを、メソッド ファイルのメソッドの `getDataWithToken` 一番上にHomeES6.jsします。

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. 次に、同じメソッドの `try` ブロックの上部に、呼び出しの上に次の行を追加します `getAccessToken`。

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>ステージングと運用に移動するときにアドインを更新する

すべての Web Officeと`localhost:44355`同様に、ステージング サーバーまたは運用サーバーに移動する準備ができたら、マニフェスト内のドメインを新しいドメインで更新する必要があります。 同様に、ドメイン ファイル内のドメインを更新web.configがあります。

ドメインはドメイン登録に表示AAD、`localhost:44355`その登録を更新して、新しいドメインが表示される場所に代って使用する必要があります。
