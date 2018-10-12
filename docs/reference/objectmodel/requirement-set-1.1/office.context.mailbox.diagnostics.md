
# <a name="diagnostics"></a><span data-ttu-id="683d2-101">診断</span><span class="sxs-lookup"><span data-stu-id="683d2-101">diagnostics</span></span>

### <span data-ttu-id="683d2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="683d2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="683d2-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="683d2-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="683d2-105">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-105">Requirements</span></span>

|<span data-ttu-id="683d2-106">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-106">Requirement</span></span>| <span data-ttu-id="683d2-107">値</span><span class="sxs-lookup"><span data-stu-id="683d2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="683d2-108">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="683d2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="683d2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="683d2-109">1.0</span></span>|
|[<span data-ttu-id="683d2-110">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="683d2-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="683d2-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="683d2-111">ReadItem</span></span>|
|[<span data-ttu-id="683d2-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="683d2-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="683d2-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="683d2-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="683d2-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="683d2-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="683d2-115">ホスト名: 文字列</span><span class="sxs-lookup"><span data-stu-id="683d2-115">hostName :String</span></span>

<span data-ttu-id="683d2-116">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="683d2-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="683d2-117">文字列は、`Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかの値になります。</span><span class="sxs-lookup"><span data-stu-id="683d2-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="683d2-118">種類:</span><span class="sxs-lookup"><span data-stu-id="683d2-118">Type:</span></span>

*   <span data-ttu-id="683d2-119">文字列</span><span class="sxs-lookup"><span data-stu-id="683d2-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="683d2-120">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-120">Requirements</span></span>

|<span data-ttu-id="683d2-121">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-121">Requirement</span></span>| <span data-ttu-id="683d2-122">値</span><span class="sxs-lookup"><span data-stu-id="683d2-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="683d2-123">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="683d2-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="683d2-124">1.0</span><span class="sxs-lookup"><span data-stu-id="683d2-124">1.0</span></span>|
|[<span data-ttu-id="683d2-125">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="683d2-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="683d2-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="683d2-126">ReadItem</span></span>|
|[<span data-ttu-id="683d2-127">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="683d2-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="683d2-128">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="683d2-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="683d2-129">hostVersion: 文字列</span><span class="sxs-lookup"><span data-stu-id="683d2-129">hostVersion :String</span></span>

<span data-ttu-id="683d2-130">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="683d2-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="683d2-p102">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="683d2-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="683d2-134">種類:</span><span class="sxs-lookup"><span data-stu-id="683d2-134">Type:</span></span>

*   <span data-ttu-id="683d2-135">文字列</span><span class="sxs-lookup"><span data-stu-id="683d2-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="683d2-136">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-136">Requirements</span></span>

|<span data-ttu-id="683d2-137">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-137">Requirement</span></span>| <span data-ttu-id="683d2-138">値</span><span class="sxs-lookup"><span data-stu-id="683d2-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="683d2-139">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="683d2-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="683d2-140">1.0</span><span class="sxs-lookup"><span data-stu-id="683d2-140">1.0</span></span>|
|[<span data-ttu-id="683d2-141">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="683d2-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="683d2-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="683d2-142">ReadItem</span></span>|
|[<span data-ttu-id="683d2-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="683d2-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="683d2-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="683d2-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="683d2-145">OWAView: 文字列</span><span class="sxs-lookup"><span data-stu-id="683d2-145">OWAView :String</span></span>

<span data-ttu-id="683d2-146">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="683d2-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="683d2-147">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="683d2-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="683d2-148">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティへのアクセスは `undefined` となります。</span><span class="sxs-lookup"><span data-stu-id="683d2-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="683d2-149">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="683d2-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="683d2-p103">`OneColumn`これは、画面の幅が狭い場合に表示されます。Outlook Web App は、このシングル コラム レイアウトを使用してスマートフォンの画面全体に表示します。</span><span class="sxs-lookup"><span data-stu-id="683d2-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="683d2-p104">`TwoColumns`画面幅がやや広い場合に表示される 。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="683d2-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="683d2-p105">`ThreeColumns`これは、画面の幅が広い場合に表示されます。Outlook Web App は、デスクトップ コンピュータの全画面表示ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="683d2-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="683d2-156">種類:</span><span class="sxs-lookup"><span data-stu-id="683d2-156">Type:</span></span>

*   <span data-ttu-id="683d2-157">文字列</span><span class="sxs-lookup"><span data-stu-id="683d2-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="683d2-158">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-158">Requirements</span></span>

|<span data-ttu-id="683d2-159">要件</span><span class="sxs-lookup"><span data-stu-id="683d2-159">Requirement</span></span>| <span data-ttu-id="683d2-160">値</span><span class="sxs-lookup"><span data-stu-id="683d2-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="683d2-161">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="683d2-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="683d2-162">1.0</span><span class="sxs-lookup"><span data-stu-id="683d2-162">1.0</span></span>|
|[<span data-ttu-id="683d2-163">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="683d2-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="683d2-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="683d2-164">ReadItem</span></span>|
|[<span data-ttu-id="683d2-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="683d2-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="683d2-166">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="683d2-166">Compose or read</span></span>|