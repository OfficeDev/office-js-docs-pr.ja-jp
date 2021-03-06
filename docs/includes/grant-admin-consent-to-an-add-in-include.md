
> [!NOTE]
> この手順が必要とされるのは、アドインを開発しているときだけです。 運用アドインが AppSource またはアプリカタログに展開されている場合、ユーザーはそのアドインを個別に信頼するか、管理者が組織のインストール時に同意することになります。

[アドインを登録](../develop/register-sso-add-in-aad-v2.md)し*た後*で、この手順を実行します。 (この手順を完了したばかりで、[ **$ADD-NAME $** ] ページの [ **API の権限**] タブがブラウザーで開かれている場合は、 **[テナント名] ボタンに [管理者の同意を許可**する] ボタンを選択して、表示される確認に対して [**はい**] を選択します。 この手順の残りの部分をスキップします。)

1. [ [Azure ポータル-アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908)] ページに移動して、アプリの登録を表示します。

1. Microsoft 365 テナントに対して***管理者***の資格情報を使用してサインインします。 たとえば、MyName@contoso.onmicrosoft.com です。

1. 表示名が **$ADD**のアプリを選択します。

1. [ **$ADD 名 $** ] ページで、[ **API アクセス許可**] を選択し、[**同意を許可**する] セクションで **[[テナント名] に対する管理者の同意を付与**する] ボタンをクリックします。 表示される確認の [**はい]** を選択します。

> [!NOTE]
> 開発者 O365 テナントを使用している場合は、この手順をベストプラクティスとしてお勧めします。 ただし、必要に応じて、開発環境で SSO アドインをサイドロードし、同意フォームをユーザーに求めることができます。 詳細については、「[サイドロード On Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) 」および「[サイドロード on Office on the web](../testing/sideload-office-add-ins-for-testing.md)」を参照してください。
