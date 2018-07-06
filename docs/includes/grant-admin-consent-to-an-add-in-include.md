
> [!NOTE]
> この手順が必要とされるのは、アドインを開発しているときだけです。実際に運用するアドインを AppSource またはアドイン カタログに展開した場合、インストール時に、各ユーザーが個別にそのアドインを信頼するか、管理者が組織のために同意することになります。

この手順は、[アドインを登録](../develop/register-sso-add-in-aad-v2.md)した*後*に実行してください。

1. 次に示す文字列内で、プレースホルダー "{application_ID}" を、アドインの登録時にコピーしたアプリケーション ID に置き換えます:  `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. そうしてできた URL をブラウザーのアドレス バーに貼り付けて、そこに移動します。

1. ダイアログが表示されたら、管理者の資格情報を使用して Office 365 テナントにサインインします。

1. その後、Microsoft Graph データにアクセスするためのアクセス許可をアドインに付与するように求めるダイアログが表示されます。**[承諾]** をクリックします。

1. ブラウザのウィンドウ/タブが、アドインの登録時に指定した **リダイレクトURL** にリダイレクトされます。 アドインの Web アプリケーションを実行している場合は、アドインのホームページがブラウザで開きます。それ以外の場合は、404 エラーが発生します。 しかし、ブラウザがホームページを開こうとしたという事実は、同意がうまく与えられたことを意味します。

>[!NOTE]
>開発者向け O365 テナントを使用している場合は、この手順をベスト プラクティスとしてお勧めします。 ただし、お好みに応じて、開発中の SSO アドインをサイドローディングして、ユーザーに同意書を要求することもできます。 詳細については、 [Windows のサイドローディング](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) および [Office Online のサイドローディング](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) を参照してください。

