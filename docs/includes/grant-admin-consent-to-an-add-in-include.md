
> [!NOTE]
> この手順が必要とされるのは、アドインを開発しているときだけです。実際に運用するアドインを AppSource またはアドイン カタログに展開した場合、インストール時に、各ユーザーが個別にそのアドインを信頼するか、管理者が組織のために同意することになります。

[アドインを登録](../develop/register-sso-add-in-aad-v2.md)し*た後*で、この手順を実行します。

1. 次に示す文字列内のプレースホルダー "{application_ID}" は、アドインの登録時にコピーしたアプリケーション ID に置き換えます: `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. そうしてできた URL をブラウザーのアドレス バーに貼り付けて、そこに移動します。

1. ダイアログが表示されたら、管理者の資格情報を使用して Office 365 テナントにサインインします。

1. その後で、Microsoft Graph データにアクセスするためのアクセス許可をアドインに付与するように求めるダイアログが表示されます。**[承諾]** をクリックします。

1. その後、ブラウザーウィンドウ/タブは、アドインの登録時に指定した**リダイレクト URL**にリダイレクトされます。 アドインの web アプリケーションが実行されている場合は、アドインのホームページがブラウザーで開きます。それ以外の場合は、404エラーが表示されます。 しかし、ブラウザーがホームページを開こうとすると、同意が正常に付与されたということです。

>[!NOTE]
>開発者 O365 テナントを使用している場合は、この手順をベストプラクティスとしてお勧めします。 ただし、必要に応じて、開発環境で SSO アドインをサイドロードし、同意フォームをユーザーに求めることができます。 詳細については、「[サイドロード on Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) 」および「[サイドロード on Office Online](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)」を参照してください。
