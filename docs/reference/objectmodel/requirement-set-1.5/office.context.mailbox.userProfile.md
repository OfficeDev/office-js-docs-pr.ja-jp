# <a name="userprofile"></a><span data-ttu-id="bbc59-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="bbc59-101">userProfile</span></span>

### <span data-ttu-id="bbc59-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="bbc59-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="bbc59-104">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-104">Requirements</span></span>

|<span data-ttu-id="bbc59-105">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-105">Requirement</span></span>| <span data-ttu-id="bbc59-106">値</span><span class="sxs-lookup"><span data-stu-id="bbc59-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbc59-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bbc59-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbc59-108">1.0</span><span class="sxs-lookup"><span data-stu-id="bbc59-108">1.0</span></span>|
|[<span data-ttu-id="bbc59-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bbc59-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bbc59-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bbc59-110">ReadItem</span></span>|
|[<span data-ttu-id="bbc59-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bbc59-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbc59-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="bbc59-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bbc59-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="bbc59-113">Members and methods</span></span>

| <span data-ttu-id="bbc59-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="bbc59-114">Member</span></span> | <span data-ttu-id="bbc59-115">種類</span><span class="sxs-lookup"><span data-stu-id="bbc59-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bbc59-116">displayName</span><span class="sxs-lookup"><span data-stu-id="bbc59-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="bbc59-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="bbc59-117">Member</span></span> |
| [<span data-ttu-id="bbc59-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="bbc59-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="bbc59-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="bbc59-119">Member</span></span> |
| [<span data-ttu-id="bbc59-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="bbc59-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="bbc59-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="bbc59-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="bbc59-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="bbc59-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="bbc59-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="bbc59-123">displayName :String</span></span>

<span data-ttu-id="bbc59-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="bbc59-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="bbc59-125">型:</span><span class="sxs-lookup"><span data-stu-id="bbc59-125">Type:</span></span>

*   <span data-ttu-id="bbc59-126">String</span><span class="sxs-lookup"><span data-stu-id="bbc59-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bbc59-127">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-127">Requirements</span></span>

|<span data-ttu-id="bbc59-128">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-128">Requirement</span></span>| <span data-ttu-id="bbc59-129">値</span><span class="sxs-lookup"><span data-stu-id="bbc59-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbc59-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bbc59-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbc59-131">1.0</span><span class="sxs-lookup"><span data-stu-id="bbc59-131">1.0</span></span>|
|[<span data-ttu-id="bbc59-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bbc59-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bbc59-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bbc59-133">ReadItem</span></span>|
|[<span data-ttu-id="bbc59-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bbc59-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbc59-135">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="bbc59-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bbc59-136">例</span><span class="sxs-lookup"><span data-stu-id="bbc59-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="bbc59-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="bbc59-137">emailAddress :String</span></span>

<span data-ttu-id="bbc59-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="bbc59-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="bbc59-139">型:</span><span class="sxs-lookup"><span data-stu-id="bbc59-139">Type:</span></span>

*   <span data-ttu-id="bbc59-140">String</span><span class="sxs-lookup"><span data-stu-id="bbc59-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bbc59-141">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-141">Requirements</span></span>

|<span data-ttu-id="bbc59-142">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-142">Requirement</span></span>| <span data-ttu-id="bbc59-143">値</span><span class="sxs-lookup"><span data-stu-id="bbc59-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbc59-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bbc59-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbc59-145">1.0</span><span class="sxs-lookup"><span data-stu-id="bbc59-145">1.0</span></span>|
|[<span data-ttu-id="bbc59-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bbc59-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bbc59-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bbc59-147">ReadItem</span></span>|
|[<span data-ttu-id="bbc59-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bbc59-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbc59-149">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="bbc59-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bbc59-150">例</span><span class="sxs-lookup"><span data-stu-id="bbc59-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="bbc59-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="bbc59-151">timeZone :String</span></span>

<span data-ttu-id="bbc59-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="bbc59-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="bbc59-153">型:</span><span class="sxs-lookup"><span data-stu-id="bbc59-153">Type:</span></span>

*   <span data-ttu-id="bbc59-154">String</span><span class="sxs-lookup"><span data-stu-id="bbc59-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bbc59-155">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-155">Requirements</span></span>

|<span data-ttu-id="bbc59-156">要件</span><span class="sxs-lookup"><span data-stu-id="bbc59-156">Requirement</span></span>| <span data-ttu-id="bbc59-157">値</span><span class="sxs-lookup"><span data-stu-id="bbc59-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbc59-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bbc59-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbc59-159">1.0</span><span class="sxs-lookup"><span data-stu-id="bbc59-159">1.0</span></span>|
|[<span data-ttu-id="bbc59-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bbc59-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bbc59-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bbc59-161">ReadItem</span></span>|
|[<span data-ttu-id="bbc59-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bbc59-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbc59-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="bbc59-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bbc59-164">例</span><span class="sxs-lookup"><span data-stu-id="bbc59-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```