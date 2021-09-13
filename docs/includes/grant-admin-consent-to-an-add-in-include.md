
> [!NOTE]
> この手順が必要とされるのは、アドインを開発しているときだけです。 実稼働アドインを AppSource またはアプリ カタログに展開すると、ユーザーは個別に信頼するか、管理者がインストール時に組織に同意します。

アドインを登録した *後* で [、この手順を実行します](../develop/register-sso-add-in-aad-v2.md)。 (この手順を完了したばかりで **、ブラウザーで $ADD-IN-NAME$** ページの **[API** アクセス許可] タブが開いている場合は、[テナント名] ボタンに対する管理者の同意を許可する] ボタンを選択し、表示される確認のために **[は** い] を選択できます。 この手順の残りの部分をスキップします。

1. Azure portal [- アプリ登録ページに移動して](https://go.microsoft.com/fwlink/?linkid=2083908) 、アプリの登録を表示します。

1. 管理者資格情報を ***使用してテナント*** にサインインMicrosoft 365します。 たとえば、MyName@contoso.onmicrosoft.com です。

1. 表示名が表示されているアプリを選択 **$ADD-IN-NAME$ を指定します**。

1. **[$ADD-IN-NAME$]** ページで **[API** のアクセス許可] を選択し、[同意の付与] セクションで、[[テナント名] の管理者の同意を許可 **する] ボタンを選択** します。 表示 **される確認に** 対して [はい] を選択します。

> [!NOTE]
> Developer O365 テナントを使用する場合は、この手順をベスト プラクティスとしてお勧めします。 ただし、必要に応じて、開発中の SSO アドインをサイドロードし、ユーザーに同意フォームを求めるメッセージを表示することもできます。 詳細については[、「Sideload on Windows」](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)と「サイドロード オン Office on the web」 を[参照してください](../testing/sideload-office-add-ins-for-testing.md)。
