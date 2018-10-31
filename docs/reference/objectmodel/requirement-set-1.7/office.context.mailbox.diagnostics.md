
# <a name="diagnostics"></a><span data-ttu-id="314dc-101">診断</span><span class="sxs-lookup"><span data-stu-id="314dc-101">diagnostics</span></span>

### <span data-ttu-id="314dc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="314dc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="314dc-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="314dc-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="314dc-105">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-105">Requirements</span></span>

|<span data-ttu-id="314dc-106">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-106">Requirement</span></span>| <span data-ttu-id="314dc-107">値</span><span class="sxs-lookup"><span data-stu-id="314dc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="314dc-108">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="314dc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="314dc-109">1.0以降</span><span class="sxs-lookup"><span data-stu-id="314dc-109">1.0</span></span>|
|[<span data-ttu-id="314dc-110">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="314dc-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="314dc-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="314dc-111">ReadItem</span></span>|
|[<span data-ttu-id="314dc-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="314dc-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="314dc-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="314dc-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="314dc-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="314dc-114">Members and methods</span></span>

| <span data-ttu-id="314dc-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="314dc-115">Member</span></span> | <span data-ttu-id="314dc-116">タイプ</span><span class="sxs-lookup"><span data-stu-id="314dc-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="314dc-117">ホスト名</span><span class="sxs-lookup"><span data-stu-id="314dc-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="314dc-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="314dc-118">Member</span></span> |
| [<span data-ttu-id="314dc-119">ホストバージョン</span><span class="sxs-lookup"><span data-stu-id="314dc-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="314dc-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="314dc-120">Member</span></span> |
| [<span data-ttu-id="314dc-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="314dc-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="314dc-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="314dc-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="314dc-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="314dc-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="314dc-124">ホスト名: 文字列</span><span class="sxs-lookup"><span data-stu-id="314dc-124">hostName :String</span></span>

<span data-ttu-id="314dc-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="314dc-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="314dc-126">文字列は、値 `Outlook`、`Mac Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="314dc-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="314dc-127">種類:</span><span class="sxs-lookup"><span data-stu-id="314dc-127">Type:</span></span>

*   <span data-ttu-id="314dc-128">文字列</span><span class="sxs-lookup"><span data-stu-id="314dc-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="314dc-129">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-129">Requirements</span></span>

|<span data-ttu-id="314dc-130">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-130">Requirement</span></span>| <span data-ttu-id="314dc-131">値</span><span class="sxs-lookup"><span data-stu-id="314dc-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="314dc-132">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="314dc-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="314dc-133">1.0以降</span><span class="sxs-lookup"><span data-stu-id="314dc-133">1.0</span></span>|
|[<span data-ttu-id="314dc-134">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="314dc-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="314dc-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="314dc-135">ReadItem</span></span>|
|[<span data-ttu-id="314dc-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="314dc-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="314dc-137">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="314dc-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="314dc-138">hostVersion: 文字列</span><span class="sxs-lookup"><span data-stu-id="314dc-138">hostVersion :String</span></span>

<span data-ttu-id="314dc-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="314dc-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="314dc-p102">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="314dc-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="314dc-143">種類:</span><span class="sxs-lookup"><span data-stu-id="314dc-143">Type:</span></span>

*   <span data-ttu-id="314dc-144">文字列</span><span class="sxs-lookup"><span data-stu-id="314dc-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="314dc-145">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-145">Requirements</span></span>

|<span data-ttu-id="314dc-146">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-146">Requirement</span></span>| <span data-ttu-id="314dc-147">値</span><span class="sxs-lookup"><span data-stu-id="314dc-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="314dc-148">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="314dc-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="314dc-149">1.0以降</span><span class="sxs-lookup"><span data-stu-id="314dc-149">1.0</span></span>|
|[<span data-ttu-id="314dc-150">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="314dc-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="314dc-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="314dc-151">ReadItem</span></span>|
|[<span data-ttu-id="314dc-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="314dc-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="314dc-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="314dc-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="314dc-154">OWAView: 文字列</span><span class="sxs-lookup"><span data-stu-id="314dc-154">OWAView :String</span></span>

<span data-ttu-id="314dc-155">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="314dc-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="314dc-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="314dc-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="314dc-157">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティへのアクセスは `undefined` となります。</span><span class="sxs-lookup"><span data-stu-id="314dc-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="314dc-158">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="314dc-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="314dc-p103">`OneColumn`これは、画面の幅が狭い場合に表示されます。Outlook Web App は、このシングル コラム レイアウトを使用してスマートフォンの画面全体に表示します。</span><span class="sxs-lookup"><span data-stu-id="314dc-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="314dc-p104">`TwoColumns`画面幅がやや広い場合に表示される 。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="314dc-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="314dc-p105">`ThreeColumns`これは、画面の幅が広い場合に表示されます。Outlook Web App は、デスクトップ コンピュータの全画面表示ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="314dc-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="314dc-165">種類:</span><span class="sxs-lookup"><span data-stu-id="314dc-165">Type:</span></span>

*   <span data-ttu-id="314dc-166">文字列</span><span class="sxs-lookup"><span data-stu-id="314dc-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="314dc-167">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-167">Requirements</span></span>

|<span data-ttu-id="314dc-168">必要条件</span><span class="sxs-lookup"><span data-stu-id="314dc-168">Requirement</span></span>| <span data-ttu-id="314dc-169">値</span><span class="sxs-lookup"><span data-stu-id="314dc-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="314dc-170">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="314dc-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="314dc-171">1.0以降</span><span class="sxs-lookup"><span data-stu-id="314dc-171">1.0</span></span>|
|[<span data-ttu-id="314dc-172">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="314dc-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="314dc-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="314dc-173">ReadItem</span></span>|
|[<span data-ttu-id="314dc-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="314dc-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="314dc-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="314dc-175">Compose or read</span></span>|