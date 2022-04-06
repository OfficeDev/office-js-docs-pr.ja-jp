---
title: シングル サインオンを使用する Node.js Office アドインを作成する
description: シングル サインオンOffice使用するNode.js ベースのアドインを作成する方法について説明します。
ms.date: 03/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: e03d023d6050f6b74ba401b1f2e0a5ed87a5cc0f
ms.sourcegitcommit: 3c5ede9c4f9782947cea07646764f76156504ff9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/06/2022
ms.locfileid: "64682246"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する Node.js Office アドインを作成する

ユーザーは、このサインイン プロセスを利用してユーザーを承認する Office および Office Web アドインにサインインできます。こうして承認されたユーザーは、アドインと Microsoft Graph への 2 度目のサインオンの必要がなくなります。概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。

この記事では、Node.js と Express を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。 ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)」を参照してください。

> [!NOTE]
> この記事で説明する手順を完了する代わりに、Yeoman ジェネレーターを使用して SSO が有効な Node.js Office アドインを作成することもできます。 Yeoman ジェネレーターは、Azure 内で SSO を構成するために必要な手順を自動化し、SSO を使用するために必要なコードを生成することで、SSO が有効なアドインの作成プロセスを簡素化します。 詳細については、「[シングル サインオン (SSO) のクイック スタート](../quickstarts/sso-quickstart.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

* [Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)

* [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

* TypeScript、バージョン 3.6.2 以降

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* コード エディター。 Visual Studio Code をお勧めします。

* Microsoft 365 サブスクリプションのOneDrive for Businessに格納されている少なくともいくつかのファイルとフォルダー。

* Microsoft Azure サブスクリプション。 このアドインには、Azure Active Directory (AD) が必要です。 Azure AD は、アプリケーションが認証および承認に使用する ID サービスを提供します。 [Microsoft Azure](https://account.windowsazure.com/SignUp) で試用版サブスクリプションを取得できます。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

1. 「[Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)」にあるリポジトリを複製するかダウンロードします。

    > [!NOTE]
    > このサンプルには、次の 3 つのバージョンがあります。
    >
    > * **Begin** フォルダーはスターター プロジェクトです。 SSO や承認に直接関連しない UI などの側面は、既に完了しています。 この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。
    > * このサンプルの **[Complete]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。 完成したバージョンを使用するには、この記事の手順に従いますが、"Begin" を "Completed" に置き換え、 **クライアント側のコードとサーバー側のコード** 化のセクション **を** スキップします。
    > * **SSOAutoSetup** バージョンは、アドインを Azure AD に登録して構成する手順の大部分を自動化する完成されたサンプルです。 SSO で動作するアドインをすばやく表示する場合には、このバージョンを使用します。 フォルダーの Readme の手順に従ってください。 Azure AD とアドインの関係をよりよく理解するために、この記事にある手動での登録およびセットアップのステップを行うことをお勧めします。

1. **[開始]** フォルダーでコマンド プロンプトを開きます。

1. コンソールで `npm install` を入力して、package.json ファイルに項目化されているすべての依存関係をインストールします。

1. コマンド`npm run install-dev-certs`を実行します。 証明書をインストールするプロンプトに対して **はい** を選択します。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Azure AD v2.0 エンドポイントにアドインを登録する

1. [Azure ポータル - アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908)ページに移動してアプリを登録します。

1. ***管理者*** 資格情報を使用してMicrosoft 365テナントにサインインします。 たとえば、MyName@contoso.onmicrosoft.com です。

1. **[新規登録]** を選択します。 **[アプリケーションを登録]** ページで、次のように値を設定します。

    * `Office-Add-in-NodeJS-SSO` に **[名前]** を設定します。
    * **[サポートされているアカウントの種類]** を **[任意の組織のディレクトリ内のアカウントと個人用の Microsoft アカウント (例: Skype、 Xbox、Outlook.com)]** に設定します。
    * アプリケーションの種類を **Web** に設定し、[ **リダイレクト URI] を [リダイレクト URI]** に `https://localhost:44355/dialog.html`設定します。
    * **[登録]** を選択します。

1. **Office-Add-in-NodeJS-SSO** ページで、**アプリケーション (クライアント) ID** と **ディレクトリ (テナント) ID** の値をコピーして保存します。 以降の手順では、それらの両方を使用します。

    > [!NOTE]
    > この **アプリケーション (クライアント) ID** は、Office クライアント アプリケーション (PowerPoint、Word、Excelなど) などの他のアプリケーションがアプリケーションへの承認されたアクセスを求める場合の "対象ユーザー" の値です。 また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. **[管理]** の下の **[認証]** を選択します。 [ **暗黙的な付与** ] セクションで、 **アクセス トークン** と **ID トークン** の両方のチェック ボックスをオンにします。 サンプルには、SSO が利用できないときに呼び出されるフォールバック認証システムがあります。 このシステムは、暗黙的フローを使用します。

1. フォームの最上部で **[保存]** を選択します。

1. **[管理]** で **[証明書とシークレット]** を選択します。 **[新しいクライアント シークレット]** ボタンを選択します。 **[説明]** に値を入力してから、**[有効期限]** に適切なオプションを選択し、**[追加]** を選択します。 *クライアント シークレットの値をすぐにコピーして、後の手順で必要になるため、先に進む前にアプリケーションIDと一緒に保存* してください。

1. **[管理]** の下の **[API の公開]** を選択します。 [ **設定** ] リンクを選択します。 これにより、"api://$App ID GUID$" という形式でアプリケーション ID URI が生成されます。ここで、$App ID GUID$ は **アプリケーション (クライアント) ID です**。

1. 生成された ID で、二重スラッシュと GUID の間に挿入 `localhost:44355/` (末尾にスラッシュ "/" が追加されていることに注意してください)。 完了したら、ID 全体にフォーム `api://localhost:44355/$App ID GUID$`が含まれている必要があります 。たとえば、次のようになります `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

1. **[Scope の追加]** ボタンをクリックします。 開いたパネルで、`access_as_user`を **[スコープ名]** として入力します。

1. **[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。

1. 管理者とユーザーの同意プロンプトを構成するためのフィールドに、現在のユーザーと同じ権限を持つアドインの Web API をOfficeクライアント アプリケーションで使用できるようにするスコープに適`access_as_user`した値を入力します。 提案:

    * **管理者の同意表示名**: Officeはユーザーとして機能できます。
    * **管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。
    * **ユーザー同意表示名**: Officeはユーザーとして機能できます。
    * **ユーザーの同意の説明**: Officeが持っているのと同じ権限を持つアドインの Web API を呼び出すようにします。

1. **[状態]** が **[有効]** に設定されていることを確認してください。

1. **[スコープの追加]** を選択します。

    > [!NOTE]
    > テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。

1. [**承認済みクライアント アプリケーション**] セクションで、次の ID を入力して、すべてのMicrosoft Officeアプリケーション エンドポイントを事前に承認します。

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`(すべてのMicrosoft Officeアプリケーション エンドポイント)

    > [!NOTE]
    > ID は`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`、次のすべてのプラットフォームでOfficeを事前に承認します。 または、何らかの理由で一部のプラットフォームでOfficeへの承認を拒否する場合は、次の ID の適切なサブセットを入力することもできます。 承認を保留するプラットフォームの ID は残しておきます。 これらのプラットフォーム上のアドインのユーザーは、Web API を呼び出すことはできませんが、アドイン内の他の機能は引き続き機能します。
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. **[クライアント アプリケーションの追加]** ボタンを選択し、表示されたパネルで [クライアント ID] をそれぞれの GUID に設定して、`api://localhost:44355/$App ID GUID$/access_as_user`のチェック ボックスをオンにします。

1. **[アプリケーションの追加]** を選択します。

1. **[管理]** の下の **[API アクセス許可]** を選択し、**[アクセス許可の追加]** を選択します。 開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。

1. アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。 以下を選択します。 アドイン自体で実際に必要なのは 1 つ目だけです。ただし、`profile`アドイン Web アプリケーションへのトークンを取得するには、Office アプリケーションに対するアクセス許可が必要です。

    * Files.Read.All
    * profile

    > [!NOTE]
    > `User.Read` アクセス許可は既定でリストされています。 必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。

1. 表示される各アクセス許可のチェック ボックスをオンにします。 アドインに必要なアクセス許可を選択したら、パネルの下部にある **[アクセス許可を追加する]** ボタンをクリックします。

1. 同じページで、**[[テナント名]に管理者の同意を与える]** ボタンを選択し、表示される確認に対して **[はい]** を選択します。

## <a name="configure-the-add-in"></a>アドインを構成する

1. コード エディターで複製プロジェクトの`\Begin`フォルダーを開きます。

1. `.ENV`ファイルを開き、以前にコピーした値を使用します。 **CLIENT_ID** を **アプリケーション (クライアント) ID** に設定し、**CLIENT_SECRET** をクライアント シークレットに設定します。 値は引用符で囲ま **ない** でください。 完了すると、ファイルは以下のようになります。

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. `\public\javascripts\fallbackAuthDialog.js`ファイルを開きます。 `msalConfig`宣言では、プレースホルダー $application_GUID here$ はアドインの登録時にコピーしたアプリケーション ID に置き換えます。 値は引用符で囲む必要があります。

1. アドイン マニフェスト ファイル "manifest\manifest_local.xml" を開き、ファイルの一番下までスクロールします。 終了タグのすぐ `</VersionOverrides>` 上に、次のマークアップがあります。

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

1. このマークアップ内の *両方の場所の* プレースホルダー "$application_GUID here$" を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。 "$" 記号は ID の一部ではないため、含めないでください。 これは、CLIENT_IDと対象ユーザーに対して使用した ID と同じです。ENV ファイル。

   > [!NOTE]
   > **リソース** 値は、アドインを登録したときに設定した **アプリケーション ID URI** です。 **[範囲]** セクションは、アドインが AppSource を通じて販売される場合に同意ダイアログ ボックスを生成するためにのみ使用されます。

## <a name="code-the-client-side"></a>クライアント側のコーディング

### <a name="create-the-sso-logic"></a>SSO ロジックを作成する

1. コード エディターで、`public\javascripts\ssoAuthES6.js`ファイルを開きます。 Internet Explorer 11 でも Promise がサポートされることを保証するコードと、アドインの唯一のボタンにハンドラーを割り当てるための`Office.onReady`呼び出しが既にあります。

   > [!NOTE]
   > 名前が示すように、ssoAuthES6.js は JavaScript ES6 構文を使用します。これは、これは、`async`と`await`の使用こそが SSO API の本質的なシンプルさを最もよく示すためです。 localhost サーバーが起動するとこのファイルは ES5 構文に変換され、サンプルが Internet Explorer 11 で実行されます。

1. Office.onReady メソッドの下に次のコードを追加します。

    > [!NOTE]
    > この記事で使用する 2 つのアクセス トークンを区別するために、getAccessToken() から返されるトークンはブートストラップ トークンと呼ばれます。 後で On-Behalf-Of フローを通じて、Microsoft Graphにアクセスできる新しいトークンと交換されます。

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exchange the bootstrap token for a new
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         Office.auth.getAccessToken call.

        }
    }
    ```

1. `TODO 1` を次のコードに置き換えます。 このコードについては、次の点に注意してください。

    * `Office.auth.getAccessToken`は、Azure AD からブートストラップ トークンを取得するよう Office に指示します。 ブートストラップ トークンは ID トークンですが、値`access-as-user`を持`scp`つ (スコープ) プロパティもあります。 このトークンは、Microsoft Graphへのアクセス許可を持つアクセス トークンに対して Web アプリケーションによって交換できます。
    * このオプションを `allowSignInPrompt` true に設定すると、ユーザーが現在Officeにサインインしていない場合、Officeポップアップ サインイン プロンプトが開きます。
    * このオプションを `allowConsentPrompt` true に設定すると、ユーザーがアドインがユーザーのAAD プロファイルにアクセスすることに同意していない場合、Office同意プロンプトが開きます。 (このプロンプトでは、Microsoft Graph スコープではなく、ユーザーのAAD プロファイルにのみ同意できます)。
    * このオプションを `forMSGraphAccess` true に設定すると、アドインはブートストラップ トークンを使用して、ID トークンとして使用するのではなく、Microsoft Graphへのアクセス許可を持つ追加のアクセス トークンを取得することをOfficeします。 テナント管理者が Microsoft Graph へのアドインのアクセスに同意していない場合、`Office.auth.getAccessToken`はエラー **13012** を返します。 アドインは、Office が Microsoft Graph スコープではなく、ユーザーの Azure AD プロファイルへの同意のみを要求できるために必要となる承認の代替システムにフォールバックすることで応答できます。 フォールバック承認システムでは、ユーザーが再度サインインする *必要があり、* ユーザーは Microsoft Graphスコープに同意するように求めることができます。 そのため`forMSGraphAccess`オプションは、同意の欠如により失敗するトークン交換をアドインが行わないことを保証します。 (前のステップで管理者の同意が与えられているため、このアドインにおいてはこのシナリオは発生しません。 ベスト プラクティスを示すことを目的として、このオプションはここに含まれています。)

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true }); 
    ```

1. `TODO 2`を以下のコードに置き換えます。 `getGraphToken`メソッドは後の手順で作成します。

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. `TODO 3`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * Microsoft 365 テナントが多要素認証を必要とするように構成されている場合は、追加の`exchangeResponse`必須要素に関する情報を`claims`含むプロパティが含まれます。 その場合は`Office.auth.getAccessToken`を再度呼び出し、`authChallenge`オプションを Claims プロパティの値に設定する必要があります。 これにより、必要なすべての認証形式をユーザーに求めるよう AAD に指示します。

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await Office.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. `TODO 4`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * `handleAADErrors`メソッドは後の手順で作成します。 Azure AD エラーは、HTTP コード 200 応答としてクライアントに返されます。 エラーがスローされないため、`catch`ブロック (`getGraphData`メソッドのもの) をトリガーしません。
    * `makeGraphApiCall`メソッドは後の手順で作成します。 これが MS Graph エンドポイントへの AJAX 呼び出しを行います。 エラーはその呼び出しの`.fail`コールバックでキャッチされます。`catch`ブロック (`getGraphData`メソッドのもの) ではありません。

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. 次に置き換えます `TODO 5` 。

    * `getAccessToken`の呼び出しからのエラーは、通常 13xxx の範囲のエラー番号を持つ`code`プロパティを持ちます。 `handleClientSideErrors`メソッドは後の手順で作成します。
    * `showMessage`メソッドは、タスク ウィンドウにテキストを表示します。

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. `getGraphData`メソッドの下に、以下の関数を追加します。 Microsoft Graph `/auth` へのアクセス許可を持つアクセス トークンに対してブートストラップ トークンとAzure ADを交換するサーバー側 Express ルートであることに注意してください。

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. `getGraphToken`メソッドの下に、以下の関数を追加します。 `error.code`は数値であり、通常は 13xxx の範囲にあることを注意してください。

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```

1. `TODO 6`を以下のコードに置き換えます。
これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to sign in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // Office.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the web.
        showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The Office.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. `TODO 7`を以下のコードに置き換えます。 これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。関数`dialogFallback`は、代替の認証システムを呼び出します。 このアドインでは、フォールバック システムはユーザーが既にログインしている場合でもユーザーのサインインを要求するダイアログを開き、msal.js および Implicit Flow を使用して Microsoft Graph へのアクセス トークンを取得します。

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. `handleClientSideErrors`関数の下に、次の関数を追加します。

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. まれに、Officeがキャッシュしたブートストラップ トークンは、Office検証時には期限が切れますが、交換のためにAzure ADに達するまでに期限切れになります。 Azure AD はエラー **AADSTS500133** で応答します。 この場合、アドインは単に再度呼び出す `getGraphData` 必要があります。 キャッシュされたブートストラップ トークンの有効期限が切れているため、Office は Azure AD から新しいものを取得します。 そのため、次のように置き換えます `TODO 8` 。

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
    {
        getGraphData();
    }
    ```

1. アドインが`getGraphData`の呼び出しの無限ループに入らないようにするため、アドインは`getGraphData`が呼び出された回数を追跡し、1 回以上再帰的に呼び出されないことを確認する必要があります。 そのため、`handleAADErrors`および`getGraphData`関数に対してグローバルなスコープにカウンター変数を作成します。 グローバル変数の適切な場所は、`Office.onReady`メソッド呼び出しのすぐ下です。

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. `if`構造 (`handleAADErrors`メソッドのもの) を次のように変更します。

    * `getGraphData`を呼び出す直前にカウンターをインクリメントします。
    * `getGraphData`が 2 回目に呼び出されていないことをテストして確認します。

    したがって、`if`構造の最終バージョンは以下のようになります。

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. 次に置き換えます `TODO 9` 。

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. ファイルを保存して閉じます。

### <a name="get-the-data-and-add-it-to-the-office-document"></a>データを取得し、Office ドキュメントへと追加する

1. `public\javascripts`フォルダーに、`data.js`という名前の新しいファイルを作成します。

1. 次の関数をファイルに追加します。 これは、Microsoft Graph へのアクセス トークンを取得したときに`getGraphData`関数によって呼び出される関数です。  

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. `TODO 10`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * このオブジェクトは、`$.ajax`メソッドのパラメーターです。
    * `/getuserdata`は、後の手順で作成するアドインのサーバー上のエクスプレス ルートです。 Microsoft Graph エンドポイントを呼び出し、その呼び出しにアクセス トークンを含めます。

    ```javascript
    {
        type: "GET",
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. `TODO11`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * `writeFileNamesToOfficeDocument`は、Graph から Office ドキュメントにデータを挿入します。 `public\javascripts\document.js`ファイルで定義されています。
    * `writeFileNamesToOfficeDocument`がエラーを返した場合、エラー メッセージは "ドキュメントにファイル名を追加できません" で始まります。

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () {
        showMessage("Your data has been added to the document.");
    })
    .catch(function (error) {
        showMessage(error);
    });
    ```

1. ファイルを保存して閉じます。

## <a name="code-the-server-side"></a>サーバー側のコーディング

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a>認証ルーターおよびトークン交換ロジックを作成する

1. ファイル`routes\authRoute.js`を開き、`require`ステートメントのすぐ下と`module.exports`ステートメントの上に以下のルート関数を追加します。 `router.get`の URL パラメーターが '/' であることにご注意ください。 このルートは URL '/auth' へのすべての HTTP リクエストを処理するルーターで定義されているため、'/auth' へのすべてのリクエストを効率的に処理します。 以前作成したクライアント側の`getGraphToken`関数が、このルートを呼び出します。  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exchange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. `TODO 12`を以下のコードに置き換えます。

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. `TODO 13` を次のコードに置き換えます。 このコードについては、次の点に注意してください。

    * これは長い`else`ブロックの始まりですが、さらにコードを追加するため、終了`}`はまだ終わりではありません。
    * `authorization`文字列は "ベアラー" の後にブートストラップ トークンが続くため、`else`ブロックの最初の行はトークンを`jwt`に割り当てています。 ("JWT" は "JSON Web Token" の略です)。
    * 2 つの`process.env.*`値は、アドインを構成したときに割り当てた定数です。
    * `requested_token_use` フォーム パラメーターは 'on_behalf_of' に設定されています。 これにより、アドインが On-Behalf-Of フロー (OBO) を使用して Microsoft Graphへのアクセス トークンを要求していることをAzure ADに伝えます。 Azure は、フォーム パラメーターに割り当てられている`assertion`ブートストラップ トークンに 、`scp``access-as-user`.
    * `scope`フォーム パラメーターは、アドインが必要とする唯一の Microsoft Graph スコープである 'Files.Read.All' に設定されます。

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. `TODO 14`を`else`ブロックを完成させる以下のコードに置き換えます。 このコードについては、次の点に注意してください。

    * const `tenant`は 'common' に設定されます。これは、アドインを Azure AD に登録したときにアドインをマルチテナントとして構成したためです。 特に **サポートされているアカウントの種類** を **任意の組織のディレクトリ内のアカウントと個人用の Microsoft アカウント (例: Skype、Xbox、Outlook.com)** に設定したときです。 アドインが登録されているのと同じMicrosoft 365テナント内のアカウントのみをサポートするように選択した場合、このコード`tenant`ではテナントの GUID に設定されます。
    * POST 要求がエラーにならない場合、Azure AD からの応答は JSON に変換され、クライアントに送信されます。 この JSON オブジェクトには、Azure AD が Microsoft Graph へのアクセス トークンを割り当てた`access_token`プロパティがあります。

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: formurlencoded(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();

            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. ファイルを保存して閉じます。

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a>Microsoft Graph からデータを取得するルートを作成する

1. プロジェクトのルートにある`app.js`ファイルを開きます。 '/dialog.html' のルートのすぐ下に、以下のルートを追加します。 このルートは、以前の手順で作成した`makeGraphApiCall`関数によって呼び出されます。

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. `TODO 15`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * このルートの呼び出し元である`makeGraphApiCall`は、Microsoft Graph へのアクセス トークンを "access_token" という名前のヘッダーとして HTTP 要求に追加しました。
    * `getGraphData`関数は`msgraph-helper.js`ファイルで定義されています。 (これは、クライアント側の`getGraphData`関数 (`ssoAuthES6.js`ファイルで定義したもの) とは異なります。)
    * `queryParamsSegment`の最後のパラメーターはハードコーディングされています。 本番環境のアドインでこのコードを再利用し、`queryParamsSegment`の一部がユーザーの入力に由来する場合、レスポンス ヘッダー インジェクション攻撃に使用できないようサニタイズされていることをご確認ください。
    * このコードは、必要なプロパティ ("name") および上位 10 のフォルダー名またはファイル名のみを指定することにより、Microsoft Graph から取得する必要があるデータを最小化します。

    ```javascript
    const graphToken = req.get('access_token');
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. `TODO 16`を以下のように置き換えます。 このコードについては、以下の点に注意してください。

    * Microsoft Graph が無効なトークンや期限切れトークンなどのエラーを返した場合、返されたオブジェクトには HTTP ステータス (401 など) に設定されたコード プロパティがあります。 コードはエラーをクライアントに中継します。 `.fail`コールバック (`makeGraphApiCall`のもの) でキャッチされます。
    * Microsoft Graph データにはアドインが必要としない OData メタデータおよび eTag が含まれているため、コードはクライアントに送信するファイル名のみを含む新しい配列を作成します。

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. ファイルを保存して閉じます。

## <a name="run-the-project"></a>プロジェクトを実行する

1. 結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。

1. `\Begin`フォルダーのルートでコマンド プロンプトを開きます。

1. コマンド`npm start`を実行します。

1. アドインを Office アプリケーション (Excel、Word、または PowerPoint) にサイドロードして、テストをする必要があります。 手順はプラットフォームによって異なります。 「[テスト用に Office アドインをサイドロードする](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)」に手順へのリンクがあります。

1. Office アプリケーションの **[ホーム]** リボンで **[アドインの表示]** ボタン (**SSO Node.js** グループ内) を選択して、作業ウィンドウ アドインを開きます。

1. **[OneDrive ファイル名の取得]** ボタンをクリックします。 Microsoft 365 Educationまたは職場のアカウント、または Microsoft アカウントでOfficeにログインしていて、SSO が期待どおりに動作している場合は、OneDrive for Businessの最初の 10 個のファイル名とフォルダー名がドキュメントに挿入されます。 (初回には 15 秒かかる場合があります)。ログインしていない場合、または SSO をサポートしていないシナリオの場合、または SSO が何らかの理由で機能していない場合は、サインインを求めるメッセージが表示されます。 サインインすると、ファイル名とフォルダー名が表示されます。

> [!NOTE]
> 以前に別の ID で Office にサインインしており、その時点で開いていた一部の Office アプリケーションがまだ開いている場合、Office が ID を変更したかのように見えても、確実に ID を変更できていない場合があります。 これが発生すると、Microsoft Graph の呼び出しが失敗するか、以前の ID のデータが返される場合があります。 これを防ぐには、必ず *他のすべての Office アプリケーションを閉じて* から、**[OneDrive ファイル名の取得]** を押してください。
