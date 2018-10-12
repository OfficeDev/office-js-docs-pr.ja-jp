
# <a name="userprofile"></a><span data-ttu-id="d56d0-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="d56d0-101">userProfile</span></span>

### <span data-ttu-id="d56d0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d56d0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d56d0-104">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-104">Requirements</span></span>

|<span data-ttu-id="d56d0-105">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-105">Requirement</span></span>| <span data-ttu-id="d56d0-106">値</span><span class="sxs-lookup"><span data-stu-id="d56d0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d56d0-107">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d56d0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d56d0-108">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d56d0-108">1.0</span></span>|
|[<span data-ttu-id="d56d0-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d56d0-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d56d0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d56d0-110">ReadItem</span></span>|
|[<span data-ttu-id="d56d0-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d56d0-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d56d0-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d56d0-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="d56d0-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="d56d0-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="d56d0-114">displayName : 文字列</span><span class="sxs-lookup"><span data-stu-id="d56d0-114">displayName :String</span></span>

<span data-ttu-id="d56d0-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d56d0-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d56d0-116">種類:</span><span class="sxs-lookup"><span data-stu-id="d56d0-116">Type:</span></span>

*   <span data-ttu-id="d56d0-117">文字列</span><span class="sxs-lookup"><span data-stu-id="d56d0-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d56d0-118">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-118">Requirements</span></span>

|<span data-ttu-id="d56d0-119">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-119">Requirement</span></span>| <span data-ttu-id="d56d0-120">値</span><span class="sxs-lookup"><span data-stu-id="d56d0-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="d56d0-121">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d56d0-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d56d0-122">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d56d0-122">1.0</span></span>|
|[<span data-ttu-id="d56d0-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d56d0-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d56d0-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d56d0-124">ReadItem</span></span>|
|[<span data-ttu-id="d56d0-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d56d0-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d56d0-126">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d56d0-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d56d0-127">例</span><span class="sxs-lookup"><span data-stu-id="d56d0-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d56d0-128">emailAddress : 文字列</span><span class="sxs-lookup"><span data-stu-id="d56d0-128">emailAddress :String</span></span>

<span data-ttu-id="d56d0-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d56d0-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d56d0-130">種類:</span><span class="sxs-lookup"><span data-stu-id="d56d0-130">Type:</span></span>

*   <span data-ttu-id="d56d0-131">文字列</span><span class="sxs-lookup"><span data-stu-id="d56d0-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d56d0-132">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-132">Requirements</span></span>

|<span data-ttu-id="d56d0-133">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-133">Requirement</span></span>| <span data-ttu-id="d56d0-134">値</span><span class="sxs-lookup"><span data-stu-id="d56d0-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="d56d0-135">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d56d0-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d56d0-136">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d56d0-136">1.0</span></span>|
|[<span data-ttu-id="d56d0-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d56d0-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d56d0-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d56d0-138">ReadItem</span></span>|
|[<span data-ttu-id="d56d0-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d56d0-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d56d0-140">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d56d0-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d56d0-141">例</span><span class="sxs-lookup"><span data-stu-id="d56d0-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d56d0-142">タイム ゾーン : 文字列</span><span class="sxs-lookup"><span data-stu-id="d56d0-142">timeZone :String</span></span>

<span data-ttu-id="d56d0-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d56d0-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d56d0-144">種類:</span><span class="sxs-lookup"><span data-stu-id="d56d0-144">Type:</span></span>

*   <span data-ttu-id="d56d0-145">文字列</span><span class="sxs-lookup"><span data-stu-id="d56d0-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d56d0-146">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-146">Requirements</span></span>

|<span data-ttu-id="d56d0-147">要件</span><span class="sxs-lookup"><span data-stu-id="d56d0-147">Requirement</span></span>| <span data-ttu-id="d56d0-148">値</span><span class="sxs-lookup"><span data-stu-id="d56d0-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="d56d0-149">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d56d0-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d56d0-150">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d56d0-150">1.0</span></span>|
|[<span data-ttu-id="d56d0-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d56d0-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d56d0-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d56d0-152">ReadItem</span></span>|
|[<span data-ttu-id="d56d0-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d56d0-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d56d0-154">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d56d0-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d56d0-155">例</span><span class="sxs-lookup"><span data-stu-id="d56d0-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```