## <a name="register-the-add-in-with-microsoft-identity-platform"></a>アドインを Microsoft ID プラットフォーム に登録する

Web サーバーを表すアプリ登録を Azure で作成する必要があります。 これにより、JavaScript のクライアント コードに対して適切なアクセス トークンを発行できるように、認証のサポートが可能になります。 この登録では、クライアントでの SSO と、Microsoft 認証ライブラリ (MSAL) を使用したフォールバック認証の両方がサポートされます。


1. Microsoft 365 テナントへの ***admin** _ 資格情報を使用して、[Azure portal](https://portal.azure.com/)にサインインします。 たとえば、_*MyName@contoso.onmicrosoft.com** です。
1. **[アプリの登録]** を選択します。 アイコンが表示されない場合は、検索バーで "アプリの登録" を検索します。

    :::image type="content" source="../images/azure-portal-select-app-registration.png" alt-text="Azure portalホーム ページ。":::

    **[アプリの登録]** ページが表示されます。

1. **[新規登録]** を選択します。

    :::image type="content" source="../images/azure-portal-select-new-registration.png" alt-text="[アプリの登録] ウィンドウでの新しい登録。":::

    [ **アプリケーションの登録] ウィンドウ** が表示されます。

1. [**管理**] で、[**アプリの登録** > **新しい登録**] を選択します。 [ **アプリケーションの登録** ] ウィンドウで、値を次のように設定します。

    * `<add-in-name>` に **[名前]** を設定します。
    * **[サポートされているアカウントの種類****] を[任意の組織のディレクトリ (任意の Azure AD ディレクトリ - マルチテナント)] と個人用 Microsoft アカウント (Skype、Xbox など)** に設定します。
    * [ **リダイレクト URI] を** プラットフォーム `<redirect-platform>` を使用するように設定し、URI を に `<redirect-uri>`設定します。

    :::image type="content" source="../images/azure-portal-register-an-application.png" alt-text="名前とサポートされているアカウントが完了したアプリケーション ウィンドウを登録します。":::

1. **[登録]** を選択します。 アプリケーション登録が作成されたことを示すメッセージが表示されます。

    :::image type="content" source="../images/azure-portal-application-created-message.png" alt-text="アプリケーションの登録が作成されたことを示すメッセージ。":::

1. **アプリケーション (クライアント) ID とディレクトリ (テナント) ID** の値をコピーして保存 **します**。 以降の手順では、それらの両方を使用します。

    :::image type="content" source="../images/azure-portal-copy-client-directory-ids.png" alt-text="クライアント ID とディレクトリ ID を表示する Contoso のアプリ登録ウィンドウ。":::

## <a name="add-a-client-secret"></a>クライアント シークレットを追加する

_アプリケーション パスワード_ と呼ばれることもあります。クライアント シークレットは、証明書の代わりにアプリが ID 自体に使用できる文字列値です。

1. [ **証明書&シークレット**] を選択します。 次に、[ **クライアント シークレット** ] タブで、[ **新しいクライアント シークレット**] を選択します。

    :::image type="content" source="../images/azure-portal-create-new-client-secret.png" alt-text="[証明書&シークレット] ウィンドウ。":::

    [ **クライアント シークレットの追加]** ウィンドウが表示されます。

1. クライアント シークレットの説明を追加します。
1. シークレットの有効期限を選択するか、カスタム有効期間を指定します。
    * クライアント シークレットの有効期間は、2 年間 (24 か月) 以下に制限されます。 24 か月を超えるカスタム有効期間を指定することはできません。
    * Microsoft では、有効期限の値を 12 か月未満に設定することをお勧めします。

    :::image type="content" source="../images/azure-portal-client-secret-description.png" alt-text="説明と有効期限が完了したクライアント シークレット ウィンドウを追加します。":::

1. **[追加]** を選択します。 新しいシークレットが作成され、値が一時的に表示されます。

> [!IMPORTANT]
> クライアント アプリケーション コードで使用する _シークレットの値を記録_ します。 このウィンドウを離れた後、このシークレット値は _再び表示されることはありません_ 。

## <a name="expose-a-web-api"></a>Web API を公開する

1. [ **API の公開] を選択します**。

    [ **API の公開** ] ウィンドウが表示されます。

    :::image type="content" source="../images/azure-portal-expose-an-api.png" alt-text="アプリ登録の [API の公開] ウィンドウ。":::

1. [ **設定]** を選択して、アプリケーション ID URI を生成します。

    :::image type="content" source="../images/azure-portal-set-api-uri.png" alt-text="アプリ登録の [API の公開] ウィンドウの [設定] ボタン。":::

    アプリケーション ID URI を設定するためのセクションは、 形式 `api://<app-id>`で生成されたアプリケーション ID URI と共に表示されます。

1. アプリケーション ID URI を に更新します `api://localhost:44355/<app-id>`。

    :::image type="content" source="../images/azure-portal-app-id-uri-details.png" alt-text="localhost ポートが 44355 に設定されている [アプリ ID URI] ペインを編集します。":::

    * **アプリケーション ID URI** には、アプリ ID (GUID) が形式 `api://<app-id>` で事前に入力されています。
    * アプリケーション ID URI 形式は次のとおりです。 `api://<fully-qualified-domain-name>/<app-id>`
    * と `<app-id>` (`fully-qualified-domain-name`GUID) の間`api://`に を挿入します。 たとえば、「 `api://contoso.com/<app-id>` 」のように入力します。
    * localhost を使用している場合、形式は である `api://localhost:<port>/<app-id>`必要があります。 たとえば、「 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7` 」のように入力します。

    その他のアプリケーション ID URI の詳細については、「 [アプリケーション マニフェスト識別子Uris 属性](/azure/active-directory/develop/reference-app-manifest#identifieruris-attribute)」を参照してください。

    > [!NOTE]
    > ドメインを所有しているにもかかわらず、そのドメインが既に所有されているというエラーが表示される場合は、「[クイック スタート: カスタム ドメイン名を Azure Active Directory に追加する](/azure/active-directory/add-custom-domain)」の手順に従って登録し、この手順を繰り返します。 (このエラーは、Microsoft 365 テナントの管理者の資格情報でサインインしていない場合にも発生する可能性があります。 手順 2 を参照してください。 サインアウトして、管理者の資格情報を使用して再度サインインし、手順 3 からプロセスを繰り返します。)

## <a name="add-a-scope"></a>スコープを追加する

1. **[スコープの追加]** を選択します。

    :::image type="content" source="../images/azure-portal-add-a-scope.png" alt-text="[スコープの追加] ボタンを選択します。":::

    [ **スコープの追加]** ウィンドウが開きます。

1. [ **スコープの追加] ウィンドウで、スコープ** の属性を指定します。

    :::image type="content" source="../images/azure-portal-add-a-scope-details.png" alt-text="値の例を含むスコープ ウィンドウを追加します。":::

    | フィールド | 説明 | 値 |
    |-------|-------------|---------|
    | **スコープ名** | スコープの名前。 一般的なスコープの名前付け規則は です `resource.operation.constraint`。 | SSO の場合、これは に設定する `access_as_user`必要があります。 |
    | **同意可能なロール** |  管理者の同意が必要かどうか、またはユーザーが管理者の承認なしで同意できるかどうかを決定します。 | SSO とサンプルを学習するには、これを **[管理者とユーザー]** に設定することをお勧めします。 <br><br>高い特権を持つアクセス許可 **の場合のみ、[管理者]** を選択します。|
    | **同意の表示名管理** | スコープの目的の簡単な説明は、管理者にのみ表示されます。 | `Read-only access to user files and profiles.` |
    | **同意の説明管理** | 管理者のみが表示するスコープによって付与されるアクセス許可の詳細な説明。 | `Allow Office to have read-only access to all user files and profiles. Office can call the app's web APIs as the current user.` |
    | **ユーザー同意表示名** | スコープの目的の簡単な説明。 **[管理者と** ユーザーに **同意できるユーザー]** を設定した場合にのみ、ユーザーに表示されます。 | `Read-only access to your files and profile.` |
    | **ユーザーの同意の説明** | スコープによって付与されるアクセス許可のより詳細な説明。 **[管理者と** ユーザーに **同意できるユーザー]** を設定した場合にのみ、ユーザーに表示されます。 | `Allow Office to have read-only access to your files and user profile.` |

1. **[状態**] を **[有効]** に設定し、[**スコープの追加**] を選択します。

    :::image type="content" source="../images/azure-portal-enable-state-add-scope-button.png" alt-text="[状態] を [有効] に設定し、[スコープの追加] ボタンを選択します。":::

    定義した新しいスコープがウィンドウに表示されます。

    :::image type="content" source="../images/azure-portal-scope-added-successfully.png" alt-text="[API の公開] ウィンドウに表示される新しいスコープ。":::

    > [!NOTE]
    > テキスト フィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、たとえば `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user` のように手順で設定された **[アプリケーション ID URI]** と自動的に一致し、最後に `/access_as_user` が追加されます。

1. [ **クライアント アプリケーションの追加] を選択します**

    :::image type="content" source="../images/azure-portal-add-a-client-application.png" alt-text="[クライアント アプリケーションの追加] を選択します。":::

    [ **クライアント アプリケーションの追加]** ウィンドウが表示されます。

1. **[クライアント ID**] に「」と入力します`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`。 この値は、すべての Microsoft Office アプリケーション エンドポイントを事前に承認します。

    > [!NOTE]
    > ID は `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` 、次のすべてのプラットフォームで Office を事前に承認します。 または、何らかの理由で一部のプラットフォームで Office への承認を拒否する場合は、次の ID の適切なサブセットを入力することもできます。 承認を保留するプラットフォームの ID は除外してください。 これらのプラットフォーム上のアドインのユーザーは Web API を呼び出せなくなりますが、アドイン内の他の機能は引き続き機能します。
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)

1. [ **承認されたスコープ**] で、チェック ボックスを `api://localhost:44355/<app-id>/access_as_user` オンにします。

1. **[アプリケーションの追加]** を選択します。

    :::image type="content" source="../images/azure-portal-add-application.png" alt-text="[クライアント アプリケーションの追加] ウィンドウ。":::

## <a name="add-microsoft-graph-permissions"></a>Microsoft Graph のアクセス許可を追加する

1. **[API アクセス許可]** を選択します。

    :::image type="content" source="../images/azure-portal-api-permissions.png" alt-text="[API アクセス許可] ウィンドウ。":::

    **[API アクセス許可**] ウィンドウが開きます。

1. [**アクセス許可を追加**] を選択します。

    :::image type="content" source="../images/azure-portal-add-a-permission.png" alt-text="[API アクセス許可] ウィンドウにアクセス許可を追加する。":::

    [ **API のアクセス許可の要求** ] ウィンドウが開きます。

1. **[Microsoft Graph]** を選択します。

    :::image type="content" source="../images/azure-portal-request-api-permissions-graph.png" alt-text="[Api のアクセス許可の要求] ウィンドウと [Microsoft Graph] ボタン。":::

1. [**委任されたアクセス許可**] を選択します。

    :::image type="content" source="../images/azure-portal-request-api-permissions-delegated.png" alt-text="委任されたアクセス許可を持つ [API のアクセス許可の要求] ウィンドウボタン。":::

1. [ **アクセス許可の選択** ] 検索ボックスで、アドインで必要なアクセス許可を検索します。 サンプルで使用される一般的な値を次に示します。

    * Files.Read
    * openid
    * profile

    > [!NOTE]
    > `User.Read` アクセス許可は既定でリストされています。 必要なアクセス許可のみを要求することをお勧めします。アドインで実際に必要ない場合は、このアクセス許可のチェック ボックスをオフにすることをお勧めします。

1. 表示される各アクセス許可のチェック ボックスをオンにします。 アクセス許可は、各アクセス許可を選択しても一覧に表示されません。 アドインに必要なアクセス許可を選択したら、[ **アクセス許可の追加**] を選択します。

    :::image type="content" source="../images/azure-portal-request-api-permissions-add-permissions.png" alt-text="一部のアクセス許可が選択されている [API のアクセス許可の要求] ウィンドウ。":::

## <a name="configure-access-token-version"></a>アクセス トークンのバージョンを構成する

アプリで許容されるアクセス トークンのバージョンを定義する必要があります。 この構成は、Azure Active Directory アプリケーション マニフェストで行われます。

### <a name="define-the-access-token-version"></a>アクセス トークンのバージョンを定義する

アクセス トークンのバージョンは、 **任意の組織ディレクトリ (任意の Azure AD ディレクトリ - マルチテナント) と個人用 Microsoft アカウント (Skype、Xbox など) のアカウント** の種類以外を選択した場合に変更される可能性があります。 次の手順を使用して、アクセス トークンのバージョンが Office SSO の使用に適していることを確認します。

1. 左側のウィンドウで **[管理]** > **[マニフェスト]** の順に選択します。

    :::image type="content" source="../images/azure-portal-manifest.png" alt-text="[Azure マニフェスト] を選択します。":::

    Azure Active Directory アプリケーション マニフェストが表示されます。

1. `accessTokenAcceptedVersion` プロパティの値として **2** を入力します。

    :::image type="content" source="../images/azure-portal-manifest-token-version.png" alt-text="受け入れられたアクセス トークンのバージョンの値。":::

1. **[保存]** を選びます。

    マニフェストが正常に更新されたことを示すメッセージがブラウザにポップアップ表示されます。

    :::image type="content" source="../images/azure-portal-manifest-updated-message.png" alt-text="マニフェストが更新されたメッセージ。":::

おめでとうございます。 アプリの登録を完了して、Office アドインの SSO を有効にしました。
