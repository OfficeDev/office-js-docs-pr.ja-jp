
# <a name="userprofile"></a><span data-ttu-id="fc917-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="fc917-101">userProfile</span></span>

### <span data-ttu-id="fc917-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="fc917-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc917-104">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-104">Requirements</span></span>

|<span data-ttu-id="fc917-105">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-105">Requirement</span></span>| <span data-ttu-id="fc917-106">値</span><span class="sxs-lookup"><span data-stu-id="fc917-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc917-107">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="fc917-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc917-108">1.0</span><span class="sxs-lookup"><span data-stu-id="fc917-108">1.0</span></span>|
|[<span data-ttu-id="fc917-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fc917-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fc917-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fc917-110">ReadItem</span></span>|
|[<span data-ttu-id="fc917-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fc917-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc917-112">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fc917-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="fc917-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="fc917-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="fc917-114">displayName : 文字列</span><span class="sxs-lookup"><span data-stu-id="fc917-114">displayName :String</span></span>

<span data-ttu-id="fc917-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="fc917-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="fc917-116">種類:</span><span class="sxs-lookup"><span data-stu-id="fc917-116">Type:</span></span>

*   <span data-ttu-id="fc917-117">文字列</span><span class="sxs-lookup"><span data-stu-id="fc917-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc917-118">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-118">Requirements</span></span>

|<span data-ttu-id="fc917-119">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-119">Requirement</span></span>| <span data-ttu-id="fc917-120">値</span><span class="sxs-lookup"><span data-stu-id="fc917-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc917-121">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="fc917-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc917-122">1.0</span><span class="sxs-lookup"><span data-stu-id="fc917-122">1.0</span></span>|
|[<span data-ttu-id="fc917-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fc917-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fc917-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fc917-124">ReadItem</span></span>|
|[<span data-ttu-id="fc917-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fc917-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc917-126">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fc917-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc917-127">例</span><span class="sxs-lookup"><span data-stu-id="fc917-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="fc917-128">emailAddress : 文字列</span><span class="sxs-lookup"><span data-stu-id="fc917-128">emailAddress :String</span></span>

<span data-ttu-id="fc917-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="fc917-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="fc917-130">型:</span><span class="sxs-lookup"><span data-stu-id="fc917-130">Type:</span></span>

*   <span data-ttu-id="fc917-131">文字列</span><span class="sxs-lookup"><span data-stu-id="fc917-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc917-132">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-132">Requirements</span></span>

|<span data-ttu-id="fc917-133">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-133">Requirement</span></span>| <span data-ttu-id="fc917-134">値</span><span class="sxs-lookup"><span data-stu-id="fc917-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc917-135">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="fc917-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc917-136">1.0</span><span class="sxs-lookup"><span data-stu-id="fc917-136">1.0</span></span>|
|[<span data-ttu-id="fc917-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fc917-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fc917-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fc917-138">ReadItem</span></span>|
|[<span data-ttu-id="fc917-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fc917-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc917-140">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fc917-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc917-141">例</span><span class="sxs-lookup"><span data-stu-id="fc917-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="fc917-142">タイム ゾーン : 文字列</span><span class="sxs-lookup"><span data-stu-id="fc917-142">timeZone :String</span></span>

<span data-ttu-id="fc917-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="fc917-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="fc917-144">種類:</span><span class="sxs-lookup"><span data-stu-id="fc917-144">Type:</span></span>

*   <span data-ttu-id="fc917-145">文字列</span><span class="sxs-lookup"><span data-stu-id="fc917-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc917-146">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-146">Requirements</span></span>

|<span data-ttu-id="fc917-147">要件</span><span class="sxs-lookup"><span data-stu-id="fc917-147">Requirement</span></span>| <span data-ttu-id="fc917-148">値</span><span class="sxs-lookup"><span data-stu-id="fc917-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc917-149">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="fc917-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc917-150">1.0</span><span class="sxs-lookup"><span data-stu-id="fc917-150">1.0</span></span>|
|[<span data-ttu-id="fc917-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fc917-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fc917-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fc917-152">ReadItem</span></span>|
|[<span data-ttu-id="fc917-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fc917-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc917-154">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fc917-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc917-155">例</span><span class="sxs-lookup"><span data-stu-id="fc917-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```