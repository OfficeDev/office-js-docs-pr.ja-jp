---
title: イベント ベースのアドインの AppSource Outlookオプション
description: イベント ベースのライセンス認証を実装する Outlookで使用できる AppSource リスト オプションについて説明します。
ms.topic: article
ms.date: 07/14/2021
localization_priority: Normal
ms.openlocfilehash: 0704b96b51841ec70aaf014924bed931c177b467
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/16/2021
ms.locfileid: "53458936"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a><span data-ttu-id="60ff9-103">イベント ベースのアドインの AppSource Outlookオプション</span><span class="sxs-lookup"><span data-stu-id="60ff9-103">AppSource listing options for your event-based Outlook add-in</span></span>

<span data-ttu-id="60ff9-104">現時点では、エンド ユーザーがイベント ベースの機能にアクセスするには、組織の管理者がアドインを展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60ff9-104">At present, add-ins must be deployed by an organization's admins for end-users to access the event-based feature functionality.</span></span> <span data-ttu-id="60ff9-105">エンド ユーザーが AppSource から直接アドインを取得した場合は、イベント ベースのライセンス認証を制限しています。</span><span class="sxs-lookup"><span data-stu-id="60ff9-105">We're restricting event-based activation if the end-user acquired the add-in directly from AppSource.</span></span> <span data-ttu-id="60ff9-106">たとえば、Contoso アドインにノードの下に少なくとも 1 つが定義された拡張ポイントが含まれる `LaunchEvent` `LaunchEvent Type` 場合 `LaunchEvents` (アドイン マニフェストの例から次の抜粋を参照してください)、アドインの自動呼び出しは、アドインが組織の管理者によってエンド ユーザーにインストールされた場合にのみ発生します。それ以外の場合、アドインの自動呼び出しはブロックされます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-106">For example, if the Contoso add-in includes the `LaunchEvent` extension point with at least one defined `LaunchEvent Type` under the `LaunchEvents` node (see the following excerpt from an example add-in manifest), the automatic invocation of the add-in only happens if the add-in was installed for the end-user by their organization's admin. Otherwise, the automatic invocation of the add-in is blocked.</span></span>

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

<span data-ttu-id="60ff9-107">エンド ユーザーまたは管理者は、AppSource または inclient ストアを介してアドインを取得できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-107">An end-user or admin can acquire add-ins through AppSource or the inclient store.</span></span> <span data-ttu-id="60ff9-108">アドインのプライマリ シナリオまたはワークフローでイベント ベースのアクティブ化が必要な場合は、管理者の展開で使用できるアドインを制限できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-108">If your add-in's primary scenario or workflow requires event-based activation, you may want to restrict your add-ins available to admin deployment.</span></span> <span data-ttu-id="60ff9-109">この制限を有効にするには、フライト コードの URL を指定できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-109">To enable that restriction, we can provide flight code URLs.</span></span> <span data-ttu-id="60ff9-110">フライト コードのおかげで、これらの特別な URL を持つエンド ユーザーだけがリストにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-110">Thanks to the flight codes, only end-users with these special URLs can access the listing.</span></span> <span data-ttu-id="60ff9-111">URL の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="60ff9-111">The following is an example URL.</span></span>

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

<span data-ttu-id="60ff9-112">ユーザーと管理者は、フライト コードが有効になっているときに、AppSource または inclient ストアの名前でアドインを明示的に検索することはできません。</span><span class="sxs-lookup"><span data-stu-id="60ff9-112">Users and admins can't explicitly search for an add-in by its name in AppSource or the inclient store when a flight code is enabled for it.</span></span> <span data-ttu-id="60ff9-113">アドインの作成者は、アドインの展開のために、これらのフライト コードを組織の管理者と非公開で共有できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-113">As the add-in creator, you can privately share these flight codes with organization admins for add-in deployment.</span></span>

> [!NOTE]
> <span data-ttu-id="60ff9-114">エンド ユーザーはフライト コードを使用してアドインをインストールすることができますが、アドインにはイベント ベースのライセンス認証は含めかねない。</span><span class="sxs-lookup"><span data-stu-id="60ff9-114">While end-users can install the add-in using a flight code, the add-in won't include event-based activation.</span></span>

## <a name="specify-a-flight-code"></a><span data-ttu-id="60ff9-115">フライト コードの指定</span><span class="sxs-lookup"><span data-stu-id="60ff9-115">Specify a flight code</span></span>

<span data-ttu-id="60ff9-116">アドインを発行するときに、その情報を Notes **for certification** で共有することで、アドインに必要なフライト コードを指定できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-116">You can specify the flight code you want for your add-in by sharing that information in the **Notes for certification** when you're publishing your add-in.</span></span> <span data-ttu-id="60ff9-117">_**重要**:_ フライト コードでは大文字と小文字が区別されます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-117">_**Important**:_ Flight codes are case-sensitive.</span></span>

![発行プロセス中の Notes の認定画面でのフライト コードの要求例を示すスクリーンショット。](../images/outlook-publish-notes-for-certification-1.png)

## <a name="deploy-add-in-with-flight-code"></a><span data-ttu-id="60ff9-119">フライト コードを使用してアドインを展開する</span><span class="sxs-lookup"><span data-stu-id="60ff9-119">Deploy add-in with flight code</span></span>

<span data-ttu-id="60ff9-120">フライト コードが設定された後、アプリ認定チームから URL を受け取る。</span><span class="sxs-lookup"><span data-stu-id="60ff9-120">After the flight codes are set, you'll receive the URL from the app certification team.</span></span> <span data-ttu-id="60ff9-121">その後、URL を管理者と非公開で共有できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-121">You can then share the URL with admins privately.</span></span>

<span data-ttu-id="60ff9-122">アドインを展開するには、管理者は次の手順を使用できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-122">To deploy the add-in, the admin can use the following steps.</span></span>

- <span data-ttu-id="60ff9-123">管理者アカウントで admin.microsoft.com または AppSource.com にサインインMicrosoft 365します。</span><span class="sxs-lookup"><span data-stu-id="60ff9-123">Sign in to admin.microsoft.com or AppSource.com with your Microsoft 365 admin account.</span></span> <span data-ttu-id="60ff9-124">アドインでシングル サインオン (SSO) が有効になっている場合は、グローバル管理者資格情報が必要です。</span><span class="sxs-lookup"><span data-stu-id="60ff9-124">If the add-in has Single sign-on (SSO) enabled, global admin credentials are needed.</span></span>
- <span data-ttu-id="60ff9-125">フライト コードの URL を Web ブラウザーに開きます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-125">Open the flight code URL into a web browser.</span></span>
- <span data-ttu-id="60ff9-126">アドインの一覧ページで、[今すぐ取得] **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="60ff9-126">On the add-in listing page, select **Get it now**.</span></span> <span data-ttu-id="60ff9-127">統合アプリ ポータルにリダイレクトする必要があります。</span><span class="sxs-lookup"><span data-stu-id="60ff9-127">You should be redirected to the integrated app portal.</span></span>

## <a name="unrestricted-appsource-listing"></a><span data-ttu-id="60ff9-128">無制限の AppSource リスト</span><span class="sxs-lookup"><span data-stu-id="60ff9-128">Unrestricted AppSource listing</span></span>

<span data-ttu-id="60ff9-129">重要なシナリオでイベント ベースのライセンス認証を使用しないアドイン (つまり、アドインが自動呼び出しなしで正常に動作する) 場合は、特別なフライト コードを使用せずに AppSource でアドインを一覧に表示する方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="60ff9-129">If your add-in doesn't use event-based activation for critical scenarios (that is, your add-in works well without automatic invocation), consider listing your add-in in AppSource without any special flight codes.</span></span> <span data-ttu-id="60ff9-130">エンド ユーザーが AppSource からアドインを取得した場合、ユーザーに対して自動ライセンス認証は行わなきます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-130">If an end-user gets your add-in from AppSource, automatic activation won't happen for the user.</span></span> <span data-ttu-id="60ff9-131">ただし、作業ウィンドウや UI レス コマンドなど、アドインの他のコンポーネントを使用できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-131">However, they can use other components of your add-in such as a task pane or UI-less command.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="60ff9-132">これは一時的な制限です。</span><span class="sxs-lookup"><span data-stu-id="60ff9-132">This is a temporary restriction.</span></span> <span data-ttu-id="60ff9-133">今後は、アドインを直接取得するエンド ユーザーに対してイベント ベースのアドインのアクティブ化を有効にする予定です。</span><span class="sxs-lookup"><span data-stu-id="60ff9-133">In future, we plan to enable event-based add-in activation for end-users who directly acquire your add-in.</span></span>

## <a name="update-existing-add-ins-to-include-event-based-activation"></a><span data-ttu-id="60ff9-134">既存のアドインを更新してイベント ベースのライセンス認証を含める</span><span class="sxs-lookup"><span data-stu-id="60ff9-134">Update existing add-ins to include event-based activation</span></span>

<span data-ttu-id="60ff9-135">既存のアドインを更新して、イベント ベースのライセンス認証を含め、検証のために再送信し、制限付きまたは無制限の AppSource リストを必要とするか決定できます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-135">You can update your existing add-in to include event-based activation then resubmit it for validation and decide if you want a restricted or unrestricted AppSource listing.</span></span>

<span data-ttu-id="60ff9-136">更新されたアドインが承認されると、既にアドインを展開している組織の管理者は、管理ポータルで更新メッセージを受信します。</span><span class="sxs-lookup"><span data-stu-id="60ff9-136">After the updated add-in has been approved, organization admins who have already deployed the add-in will receive an update message in the admin portal.</span></span> <span data-ttu-id="60ff9-137">メッセージは、イベント ベースのライセンス認証の変更について管理者にアドバイスします。</span><span class="sxs-lookup"><span data-stu-id="60ff9-137">The message advises the admin about the event-based activation changes.</span></span> <span data-ttu-id="60ff9-138">管理者が変更を承諾すると、更新プログラムはエンド ユーザーに展開されます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-138">After the admin accepts the changes, the update will be deployed to end-users.</span></span>

![[統合されたアプリ] 画面のアプリ更新通知のスクリーンショット。](../images/outlook-deploy-update-notification.png)

<span data-ttu-id="60ff9-140">アドインを独自にインストールしたエンド ユーザーの場合、イベント ベースのアクティブ化機能は、アドインが更新された後でも機能しません。</span><span class="sxs-lookup"><span data-stu-id="60ff9-140">For end-users who installed the add-in on their own, the event-based activation feature won't work even after the add-in has been updated.</span></span>

## <a name="admin-consent-for-installing-event-based-add-ins"></a><span data-ttu-id="60ff9-141">イベント ベースのアドインをインストールする管理者の同意</span><span class="sxs-lookup"><span data-stu-id="60ff9-141">Admin consent for installing event-based add-ins</span></span>

<span data-ttu-id="60ff9-142">管理センターの [統合アプリ] セクションからイベント ベースのアドインが展開されるたびに、管理者は展開ウィザードでアドインのイベント ベースのアクティブ化機能に関する詳細を取得します。</span><span class="sxs-lookup"><span data-stu-id="60ff9-142">Whenever an event-based add-in is deployed from the **Integrated Apps** section of the admin center, the admin gets details about the add-in's event-based activation capabilities in the deployment wizard.</span></span> <span data-ttu-id="60ff9-143">詳細は、[アプリのアクセス **許可と機能] セクションに表示** されます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-143">The details appear in the **App Permissions and Capabilities** section.</span></span> <span data-ttu-id="60ff9-144">管理者は、アドインが自動的にアクティブ化できるすべてのイベントを表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60ff9-144">The admin should see all the events where the add-in can automatically activate.</span></span>

![新しいアプリを展開するときに、[アクセス許可の要求を受け入れる] 画面のスクリーンショット。](../images/outlook-deploy-accept-permissions-requests.png)

<span data-ttu-id="60ff9-146">同様に、既存のアドインがイベント ベースの機能に更新された場合、管理者はアドインに "Update Pending" 状態を表示します。</span><span class="sxs-lookup"><span data-stu-id="60ff9-146">Similarly, when an existing add-in is updated to event-based functionality, the admin sees an "Update Pending" status on the add-in.</span></span> <span data-ttu-id="60ff9-147">更新されたアドインは、アドインが自動的にアクティブ化できる一連のイベントを含む、[アプリのアクセス許可と機能] セクションに示されている変更に管理者が同意した場合にのみ展開されます。</span><span class="sxs-lookup"><span data-stu-id="60ff9-147">The updated add-in is deployed only if the admin consents to the changes noted in the **App Permissions and Capabilities** section, including the set of events where the add-in can automatically activate.</span></span>

<span data-ttu-id="60ff9-148">アドインに新しい情報を追加する度に、管理者は管理ポータルに更新フローを表示し、追加のイベントに同意 `LaunchEvent Type` する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60ff9-148">Each time you add any new `LaunchEvent Type` to your add-in, admins will see the update flow in the admin portal and need to provide consent for additional events.</span></span>

![更新されたアプリを展開する場合の "更新" フローのスクリーンショット。](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a><span data-ttu-id="60ff9-150">関連項目</span><span class="sxs-lookup"><span data-stu-id="60ff9-150">See also</span></span>

- [<span data-ttu-id="60ff9-151">イベント ベースのOutlook用にアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="60ff9-151">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
