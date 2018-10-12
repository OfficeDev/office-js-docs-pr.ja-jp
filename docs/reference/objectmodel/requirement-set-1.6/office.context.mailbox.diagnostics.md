
# <a name="diagnostics"></a><span data-ttu-id="e75a6-101">診断</span><span class="sxs-lookup"><span data-stu-id="e75a6-101">diagnostics</span></span>

### <span data-ttu-id="e75a6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="e75a6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="e75a6-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e75a6-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e75a6-105">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-105">Requirements</span></span>

|<span data-ttu-id="e75a6-106">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-106">Requirement</span></span>| <span data-ttu-id="e75a6-107">値</span><span class="sxs-lookup"><span data-stu-id="e75a6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e75a6-108">メールボックスの要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="e75a6-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e75a6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e75a6-109">1.0</span></span>|
|[<span data-ttu-id="e75a6-110">アクセス許可の最小レベル</span><span class="sxs-lookup"><span data-stu-id="e75a6-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e75a6-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e75a6-111">ReadItem</span></span>|
|[<span data-ttu-id="e75a6-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e75a6-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e75a6-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e75a6-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e75a6-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e75a6-114">Members and methods</span></span>

| <span data-ttu-id="e75a6-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="e75a6-115">Member</span></span> | <span data-ttu-id="e75a6-116">型</span><span class="sxs-lookup"><span data-stu-id="e75a6-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e75a6-117">hostname</span><span class="sxs-lookup"><span data-stu-id="e75a6-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="e75a6-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e75a6-118">Member</span></span> |
| [<span data-ttu-id="e75a6-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="e75a6-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="e75a6-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="e75a6-120">Member</span></span> |
| [<span data-ttu-id="e75a6-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="e75a6-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="e75a6-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="e75a6-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e75a6-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="e75a6-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="e75a6-124">ホスト名: 文字列</span><span class="sxs-lookup"><span data-stu-id="e75a6-124">hostName :String</span></span>

<span data-ttu-id="e75a6-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e75a6-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="e75a6-126">文字列は、 `Outlook`、`Mac Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかの値になります。</span><span class="sxs-lookup"><span data-stu-id="e75a6-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="e75a6-127">型:</span><span class="sxs-lookup"><span data-stu-id="e75a6-127">Type:</span></span>

*   <span data-ttu-id="e75a6-128">文字列</span><span class="sxs-lookup"><span data-stu-id="e75a6-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e75a6-129">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-129">Requirements</span></span>

|<span data-ttu-id="e75a6-130">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-130">Requirement</span></span>| <span data-ttu-id="e75a6-131">値</span><span class="sxs-lookup"><span data-stu-id="e75a6-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="e75a6-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e75a6-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e75a6-133">1.0</span><span class="sxs-lookup"><span data-stu-id="e75a6-133">1.0</span></span>|
|[<span data-ttu-id="e75a6-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e75a6-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e75a6-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e75a6-135">ReadItem</span></span>|
|[<span data-ttu-id="e75a6-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e75a6-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e75a6-137">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e75a6-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="e75a6-138">hostVersion: 文字列</span><span class="sxs-lookup"><span data-stu-id="e75a6-138">hostVersion :String</span></span>

<span data-ttu-id="e75a6-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e75a6-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="e75a6-p102">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="e75a6-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="e75a6-143">型:</span><span class="sxs-lookup"><span data-stu-id="e75a6-143">Type:</span></span>

*   <span data-ttu-id="e75a6-144">文字列</span><span class="sxs-lookup"><span data-stu-id="e75a6-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e75a6-145">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-145">Requirements</span></span>

|<span data-ttu-id="e75a6-146">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-146">Requirement</span></span>| <span data-ttu-id="e75a6-147">値</span><span class="sxs-lookup"><span data-stu-id="e75a6-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="e75a6-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e75a6-148">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e75a6-149">1.0</span><span class="sxs-lookup"><span data-stu-id="e75a6-149">1.0</span></span>|
|[<span data-ttu-id="e75a6-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e75a6-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e75a6-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e75a6-151">ReadItem</span></span>|
|[<span data-ttu-id="e75a6-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e75a6-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e75a6-153">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e75a6-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="e75a6-154">OWAView: 文字列</span><span class="sxs-lookup"><span data-stu-id="e75a6-154">OWAView :String</span></span>

<span data-ttu-id="e75a6-155">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e75a6-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="e75a6-156">返される文字列は、 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかの値になります。</span><span class="sxs-lookup"><span data-stu-id="e75a6-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="e75a6-157">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="e75a6-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="e75a6-158">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="e75a6-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="e75a6-p103">`OneColumn`画面幅が狭い場合に表示される 。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="e75a6-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="e75a6-p104">`TwoColumns`画面幅がやや広い場合に表示される 。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="e75a6-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="e75a6-p105">`ThreeColumns`画面幅が広い場合に表示される 。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="e75a6-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="e75a6-165">型:</span><span class="sxs-lookup"><span data-stu-id="e75a6-165">Type:</span></span>

*   <span data-ttu-id="e75a6-166">文字列</span><span class="sxs-lookup"><span data-stu-id="e75a6-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e75a6-167">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-167">Requirements</span></span>

|<span data-ttu-id="e75a6-168">要件</span><span class="sxs-lookup"><span data-stu-id="e75a6-168">Requirement</span></span>| <span data-ttu-id="e75a6-169">値</span><span class="sxs-lookup"><span data-stu-id="e75a6-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="e75a6-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e75a6-170">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e75a6-171">1.0</span><span class="sxs-lookup"><span data-stu-id="e75a6-171">1.0</span></span>|
|[<span data-ttu-id="e75a6-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e75a6-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e75a6-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e75a6-173">ReadItem</span></span>|
|[<span data-ttu-id="e75a6-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e75a6-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e75a6-175">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e75a6-175">Compose or read</span></span>|