
# <a name="userprofile"></a><span data-ttu-id="33fe7-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="33fe7-101">userProfile</span></span>

### <span data-ttu-id="33fe7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="33fe7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="33fe7-104">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-104">Requirements</span></span>

|<span data-ttu-id="33fe7-105">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-105">Requirement</span></span>| <span data-ttu-id="33fe7-106">値</span><span class="sxs-lookup"><span data-stu-id="33fe7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="33fe7-107">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="33fe7-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33fe7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="33fe7-108">1.0</span></span>|
|[<span data-ttu-id="33fe7-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="33fe7-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33fe7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33fe7-110">ReadItem</span></span>|
|[<span data-ttu-id="33fe7-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="33fe7-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="33fe7-112">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="33fe7-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="33fe7-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="33fe7-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="33fe7-114">displayName : 文字列</span><span class="sxs-lookup"><span data-stu-id="33fe7-114">displayName :String</span></span>

<span data-ttu-id="33fe7-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="33fe7-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="33fe7-116">種類:</span><span class="sxs-lookup"><span data-stu-id="33fe7-116">Type:</span></span>

*   <span data-ttu-id="33fe7-117">文字列</span><span class="sxs-lookup"><span data-stu-id="33fe7-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33fe7-118">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-118">Requirements</span></span>

|<span data-ttu-id="33fe7-119">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-119">Requirement</span></span>| <span data-ttu-id="33fe7-120">値</span><span class="sxs-lookup"><span data-stu-id="33fe7-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="33fe7-121">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="33fe7-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33fe7-122">1.0</span><span class="sxs-lookup"><span data-stu-id="33fe7-122">1.0</span></span>|
|[<span data-ttu-id="33fe7-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="33fe7-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33fe7-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33fe7-124">ReadItem</span></span>|
|[<span data-ttu-id="33fe7-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="33fe7-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="33fe7-126">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="33fe7-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="33fe7-127">例</span><span class="sxs-lookup"><span data-stu-id="33fe7-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="33fe7-128">emailAddress : 文字列</span><span class="sxs-lookup"><span data-stu-id="33fe7-128">emailAddress :String</span></span>

<span data-ttu-id="33fe7-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="33fe7-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="33fe7-130">型:</span><span class="sxs-lookup"><span data-stu-id="33fe7-130">Type:</span></span>

*   <span data-ttu-id="33fe7-131">文字列</span><span class="sxs-lookup"><span data-stu-id="33fe7-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33fe7-132">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-132">Requirements</span></span>

|<span data-ttu-id="33fe7-133">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-133">Requirement</span></span>| <span data-ttu-id="33fe7-134">値</span><span class="sxs-lookup"><span data-stu-id="33fe7-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="33fe7-135">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="33fe7-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33fe7-136">1.0</span><span class="sxs-lookup"><span data-stu-id="33fe7-136">1.0</span></span>|
|[<span data-ttu-id="33fe7-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="33fe7-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33fe7-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33fe7-138">ReadItem</span></span>|
|[<span data-ttu-id="33fe7-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="33fe7-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="33fe7-140">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="33fe7-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="33fe7-141">例</span><span class="sxs-lookup"><span data-stu-id="33fe7-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="33fe7-142">timeZone :文字列</span><span class="sxs-lookup"><span data-stu-id="33fe7-142">timeZone :String</span></span>

<span data-ttu-id="33fe7-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="33fe7-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="33fe7-144">種類:</span><span class="sxs-lookup"><span data-stu-id="33fe7-144">Type:</span></span>

*   <span data-ttu-id="33fe7-145">文字列</span><span class="sxs-lookup"><span data-stu-id="33fe7-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33fe7-146">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-146">Requirements</span></span>

|<span data-ttu-id="33fe7-147">要件</span><span class="sxs-lookup"><span data-stu-id="33fe7-147">Requirement</span></span>| <span data-ttu-id="33fe7-148">値</span><span class="sxs-lookup"><span data-stu-id="33fe7-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="33fe7-149">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="33fe7-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33fe7-150">1.0</span><span class="sxs-lookup"><span data-stu-id="33fe7-150">1.0</span></span>|
|[<span data-ttu-id="33fe7-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="33fe7-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33fe7-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33fe7-152">ReadItem</span></span>|
|[<span data-ttu-id="33fe7-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="33fe7-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="33fe7-154">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="33fe7-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="33fe7-155">例</span><span class="sxs-lookup"><span data-stu-id="33fe7-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```