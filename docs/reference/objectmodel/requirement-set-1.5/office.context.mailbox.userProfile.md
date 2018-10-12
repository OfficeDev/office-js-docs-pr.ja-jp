# <a name="userprofile"></a><span data-ttu-id="cef5c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="cef5c-101">userProfile</span></span>

### <span data-ttu-id="cef5c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="cef5c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="cef5c-104">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-104">Requirements</span></span>

|<span data-ttu-id="cef5c-105">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-105">Requirement</span></span>| <span data-ttu-id="cef5c-106">値</span><span class="sxs-lookup"><span data-stu-id="cef5c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="cef5c-107">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="cef5c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cef5c-108">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cef5c-108">1.0</span></span>|
|[<span data-ttu-id="cef5c-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cef5c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cef5c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cef5c-110">ReadItem</span></span>|
|[<span data-ttu-id="cef5c-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cef5c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cef5c-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cef5c-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cef5c-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="cef5c-113">Members and methods</span></span>

| <span data-ttu-id="cef5c-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="cef5c-114">Member</span></span> | <span data-ttu-id="cef5c-115">種類</span><span class="sxs-lookup"><span data-stu-id="cef5c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cef5c-116">displayName</span><span class="sxs-lookup"><span data-stu-id="cef5c-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="cef5c-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="cef5c-117">Member</span></span> |
| [<span data-ttu-id="cef5c-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="cef5c-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="cef5c-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="cef5c-119">Member</span></span> |
| [<span data-ttu-id="cef5c-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="cef5c-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="cef5c-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="cef5c-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="cef5c-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="cef5c-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="cef5c-123">displayName : 文字列</span><span class="sxs-lookup"><span data-stu-id="cef5c-123">displayName :String</span></span>

<span data-ttu-id="cef5c-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="cef5c-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="cef5c-125">種類:</span><span class="sxs-lookup"><span data-stu-id="cef5c-125">Type:</span></span>

*   <span data-ttu-id="cef5c-126">文字列</span><span class="sxs-lookup"><span data-stu-id="cef5c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cef5c-127">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-127">Requirements</span></span>

|<span data-ttu-id="cef5c-128">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-128">Requirement</span></span>| <span data-ttu-id="cef5c-129">値</span><span class="sxs-lookup"><span data-stu-id="cef5c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="cef5c-130">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="cef5c-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cef5c-131">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cef5c-131">1.0</span></span>|
|[<span data-ttu-id="cef5c-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cef5c-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cef5c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cef5c-133">ReadItem</span></span>|
|[<span data-ttu-id="cef5c-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cef5c-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cef5c-135">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cef5c-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cef5c-136">例</span><span class="sxs-lookup"><span data-stu-id="cef5c-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="cef5c-137">emailAddress : 文字列</span><span class="sxs-lookup"><span data-stu-id="cef5c-137">emailAddress :String</span></span>

<span data-ttu-id="cef5c-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="cef5c-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="cef5c-139">種類:</span><span class="sxs-lookup"><span data-stu-id="cef5c-139">Type:</span></span>

*   <span data-ttu-id="cef5c-140">文字列</span><span class="sxs-lookup"><span data-stu-id="cef5c-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cef5c-141">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-141">Requirements</span></span>

|<span data-ttu-id="cef5c-142">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-142">Requirement</span></span>| <span data-ttu-id="cef5c-143">値</span><span class="sxs-lookup"><span data-stu-id="cef5c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="cef5c-144">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="cef5c-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cef5c-145">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cef5c-145">1.0</span></span>|
|[<span data-ttu-id="cef5c-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cef5c-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cef5c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cef5c-147">ReadItem</span></span>|
|[<span data-ttu-id="cef5c-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cef5c-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cef5c-149">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cef5c-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cef5c-150">例</span><span class="sxs-lookup"><span data-stu-id="cef5c-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="cef5c-151">タイム ゾーン : 文字列</span><span class="sxs-lookup"><span data-stu-id="cef5c-151">timeZone :String</span></span>

<span data-ttu-id="cef5c-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="cef5c-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="cef5c-153">種類:</span><span class="sxs-lookup"><span data-stu-id="cef5c-153">Type:</span></span>

*   <span data-ttu-id="cef5c-154">文字列</span><span class="sxs-lookup"><span data-stu-id="cef5c-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cef5c-155">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-155">Requirements</span></span>

|<span data-ttu-id="cef5c-156">要件</span><span class="sxs-lookup"><span data-stu-id="cef5c-156">Requirement</span></span>| <span data-ttu-id="cef5c-157">値</span><span class="sxs-lookup"><span data-stu-id="cef5c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="cef5c-158">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="cef5c-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cef5c-159">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cef5c-159">1.0</span></span>|
|[<span data-ttu-id="cef5c-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cef5c-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cef5c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cef5c-161">ReadItem</span></span>|
|[<span data-ttu-id="cef5c-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cef5c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cef5c-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cef5c-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cef5c-164">例</span><span class="sxs-lookup"><span data-stu-id="cef5c-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```