

1. <span data-ttu-id="e4037-101">[https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com) に移動します。</span><span class="sxs-lookup"><span data-stu-id="e4037-101">Navigate to [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).</span></span>

1. <span data-ttu-id="e4037-102">***管理者***の資格情報を使用して Office 365 テナントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="e4037-102">Sign-in with the admin credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="e4037-103">たとえば、MyName@contoso.onmicrosoft.com</span><span class="sxs-lookup"><span data-stu-id="e4037-103">For example, MyName@contoso.onmicrosoft.com</span></span>

1. <span data-ttu-id="e4037-104">**[アプリの追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4037-104">Click **Add an app**.</span></span>

1. <span data-ttu-id="e4037-105">ダイアログが表示されたら、アプリ名として **$ADD-IN-NAME$** を入力して、**[アプリケーションの作成]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4037-105">When prompted, use “Office-Add-in-ASPNET-SSO” as the app name, and then press Create application.</span></span>

1. <span data-ttu-id="e4037-p102">アプリの構成ページが開いたら、**[アプリケーション ID]** をコピーして保存します。これは、この後の手順で使用します。</span><span class="sxs-lookup"><span data-stu-id="e4037-p102">When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e4037-p103">この ID は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。</span><span class="sxs-lookup"><span data-stu-id="e4037-p103">This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="e4037-p104">**[アプリケーション シークレット]** セクションで、**[新しいパスワードを生成する]** をクリックします。新しいパスワード (「アプリケーション シークレット」とも呼びます) が示されたポップアップ ダイアログが開きます。*このパスワードをすぐにコピーして、アプリケーション ID と共に保存します。* これは、この後の手順で必要になります。その後で、ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e4037-p104">In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.</span></span>

1. <span data-ttu-id="e4037-115">**[プラットフォーム]** セクションで、**[プラットフォームの追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4037-115">In the **Platforms** section, click **Add Platform**.</span></span>

1. <span data-ttu-id="e4037-116">開いたダイアログで、**[Web API]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e4037-116">In the dialog that opens, select **Web API**.</span></span>

1. <span data-ttu-id="e4037-117">**[アプリケーション ID URI]** が、"api://$App ID GUID$" という形式で生成されています。</span><span class="sxs-lookup"><span data-stu-id="e4037-117">An **Application ID URI** has been generated of the form “api://{App ID GUID}”.</span></span> <span data-ttu-id="e4037-118">ダブル スラッシュと GUID の間に **$FQDN-WITHOUT-PROTOCOL$** (末尾にスラッシュ "/" を付けて) を挿入します。</span><span class="sxs-lookup"><span data-stu-id="e4037-118">Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="e4037-119">全体の ID は `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$` の形式になります。例: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`。</span><span class="sxs-lookup"><span data-stu-id="e4037-119">The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e4037-120">ドメインを所有しているにもかかわらず、そのドメインが既に所有されているというエラーが表示される場合は、「[クイック スタート: カスタム ドメイン名を Azure Active Directory に追加する](/azure/active-directory/add-custom-domain)」の手順に従って登録し、この手順を繰り返します。</span><span class="sxs-lookup"><span data-stu-id="e4037-120">If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step.</span></span> <span data-ttu-id="e4037-121">(このエラーは、Office 365 テナントで管理者の資格情報を使用してサインインしていない場合にも発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="e4037-121">(This error can also occur if you are not signed in with credentials of an admin in the Office 365 tenancy.</span></span> <span data-ttu-id="e4037-122">手順 2 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4037-122">See step 2.</span></span> <span data-ttu-id="e4037-123">サインアウトして、管理者の資格情報を使用して再度サインインし、手順 3 からプロセスを繰り返します。)</span><span class="sxs-lookup"><span data-stu-id="e4037-123">Sign out and sign in again with admin credentials and repeat the process from step 3.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="e4037-124">**[アプリケーション ID URI]** のすぐ下の **[スコープ]** 名のドメイン部分は、一致するように自動的に変更され、末尾に`/access_as_user` が追加されます。例: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="e4037-124">The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match.</span></span>

1. <span data-ttu-id="e4037-p107">**[事前承認済みアプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。 次のそれぞれの ID を事前承認する必要があります。 1 つの ID を入力するたびに、新しい空のテキスト ボックスが表示されます。 (GUID のみを入力してください。)</span><span class="sxs-lookup"><span data-stu-id="e4037-p107">In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized. Each time you enter one, a new empty textbox appears. (Enter only the GUID.)</span></span>
    * <span data-ttu-id="e4037-129">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="e4037-129">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="e4037-130">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="e4037-130">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span></span>
    * <span data-ttu-id="e4037-131">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="e4037-131">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span></span>

1. <span data-ttu-id="e4037-132">それぞれの **[アプリケーション ID]** の横の **[スコープ]** ドロップダウンを開いて、`api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user` のボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="e4037-132">Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span></span>

1. <span data-ttu-id="e4037-133">**[プラットフォーム]** セクションの上部にある **[プラットフォームの追加]** を再度クリックして、**[Web]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e4037-133">Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.</span></span>

1. <span data-ttu-id="e4037-134">**[プラットフォーム]** の下側の新しい **[Web]** セクションで、**[リダイレクト URL]** として `https://$FQDN-WITHOUT-PROTOCOL$` を入力します。</span><span class="sxs-lookup"><span data-stu-id="e4037-134">In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span></span>

1. <span data-ttu-id="e4037-p108">**[Microsoft Graph のアクセス許可]** セクションを下にスクロールして、**[委任されたアクセス許可]** サブセクションを表示します。**[追加]** ボタンを使用して、**[アクセス許可の選択]** ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="e4037-p108">Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.</span></span>

1. <span data-ttu-id="e4037-137">ダイアログ ボックスで、`profile` のチェック ボックスと、アドインに必要なその他の AAD と Microsoft Graph のアクセス許可をオンにします。</span><span class="sxs-lookup"><span data-stu-id="e4037-137">In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs.</span></span> <span data-ttu-id="e4037-138">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="e4037-138">The following are examples:</span></span>

    * <span data-ttu-id="e4037-139">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="e4037-139">Files.Read.All</span></span>
    * <span data-ttu-id="e4037-140">offline_access</span><span class="sxs-lookup"><span data-stu-id="e4037-140">offline_access</span></span>
    * <span data-ttu-id="e4037-141">openid</span><span class="sxs-lookup"><span data-stu-id="e4037-141">openid</span></span>
    * <span data-ttu-id="e4037-142">profile</span><span class="sxs-lookup"><span data-stu-id="e4037-142">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="e4037-143">`User.Read` アクセス許可は既定でリストされています。</span><span class="sxs-lookup"><span data-stu-id="e4037-143">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="e4037-144">必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e4037-144">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.</span></span>

1. <span data-ttu-id="e4037-145">ダイアログの下部にある **[OK]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4037-145">At the bottom of the dialog, click **OK**.</span></span>

1. <span data-ttu-id="e4037-146">登録ページの下部にある **[保存]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e4037-146">At the bottom of the registration page, click **Save**.</span></span>
