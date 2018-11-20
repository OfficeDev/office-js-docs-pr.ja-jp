
# <a name="userprofile"></a><span data-ttu-id="f9bff-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="f9bff-101">userProfile</span></span>

### <span data-ttu-id="f9bff-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="f9bff-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9bff-104">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-104">Requirements</span></span>

|<span data-ttu-id="f9bff-105">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-105">Requirement</span></span>| <span data-ttu-id="f9bff-106">値</span><span class="sxs-lookup"><span data-stu-id="f9bff-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9bff-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f9bff-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9bff-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f9bff-108">1.0</span></span>|
|[<span data-ttu-id="f9bff-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f9bff-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9bff-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9bff-110">ReadItem</span></span>|
|[<span data-ttu-id="f9bff-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f9bff-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9bff-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f9bff-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f9bff-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f9bff-113">Members and methods</span></span>

| <span data-ttu-id="f9bff-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="f9bff-114">Member</span></span> | <span data-ttu-id="f9bff-115">種類</span><span class="sxs-lookup"><span data-stu-id="f9bff-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f9bff-116">accountType</span><span class="sxs-lookup"><span data-stu-id="f9bff-116">AccountType</span></span>](#accounttype-string) | <span data-ttu-id="f9bff-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="f9bff-117">Member</span></span> |
| [<span data-ttu-id="f9bff-118">displayName</span><span class="sxs-lookup"><span data-stu-id="f9bff-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="f9bff-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="f9bff-119">Member</span></span> |
| [<span data-ttu-id="f9bff-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f9bff-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="f9bff-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f9bff-121">Member</span></span> |
| [<span data-ttu-id="f9bff-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="f9bff-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="f9bff-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="f9bff-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f9bff-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="f9bff-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="f9bff-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="f9bff-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="f9bff-126">現在、このメンバーは Outlook 2016 for Mac 以降 (ビルド 16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f9bff-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="f9bff-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="f9bff-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="f9bff-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="f9bff-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="f9bff-129">値</span><span class="sxs-lookup"><span data-stu-id="f9bff-129">Value</span></span> | <span data-ttu-id="f9bff-130">説明</span><span class="sxs-lookup"><span data-stu-id="f9bff-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="f9bff-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="f9bff-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="f9bff-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="f9bff-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="f9bff-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="f9bff-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="f9bff-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="f9bff-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="f9bff-135">種類:</span><span class="sxs-lookup"><span data-stu-id="f9bff-135">Type:</span></span>

*   <span data-ttu-id="f9bff-136">String</span><span class="sxs-lookup"><span data-stu-id="f9bff-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9bff-137">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-137">Requirements</span></span>

|<span data-ttu-id="f9bff-138">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-138">Requirement</span></span>| <span data-ttu-id="f9bff-139">値</span><span class="sxs-lookup"><span data-stu-id="f9bff-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9bff-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f9bff-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9bff-141">1.6</span><span class="sxs-lookup"><span data-stu-id="f9bff-141">-16</span></span> |
|[<span data-ttu-id="f9bff-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f9bff-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9bff-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9bff-143">ReadItem</span></span>|
|[<span data-ttu-id="f9bff-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f9bff-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9bff-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f9bff-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9bff-146">例</span><span class="sxs-lookup"><span data-stu-id="f9bff-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="f9bff-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="f9bff-147">displayName :String</span></span>

<span data-ttu-id="f9bff-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f9bff-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="f9bff-149">型:</span><span class="sxs-lookup"><span data-stu-id="f9bff-149">Type:</span></span>

*   <span data-ttu-id="f9bff-150">String</span><span class="sxs-lookup"><span data-stu-id="f9bff-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9bff-151">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-151">Requirements</span></span>

|<span data-ttu-id="f9bff-152">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-152">Requirement</span></span>| <span data-ttu-id="f9bff-153">値</span><span class="sxs-lookup"><span data-stu-id="f9bff-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9bff-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f9bff-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9bff-155">1.0</span><span class="sxs-lookup"><span data-stu-id="f9bff-155">1.0</span></span>|
|[<span data-ttu-id="f9bff-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f9bff-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9bff-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9bff-157">ReadItem</span></span>|
|[<span data-ttu-id="f9bff-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f9bff-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9bff-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f9bff-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9bff-160">例</span><span class="sxs-lookup"><span data-stu-id="f9bff-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="f9bff-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="f9bff-161">emailAddress :String</span></span>

<span data-ttu-id="f9bff-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f9bff-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="f9bff-163">型:</span><span class="sxs-lookup"><span data-stu-id="f9bff-163">Type:</span></span>

*   <span data-ttu-id="f9bff-164">String</span><span class="sxs-lookup"><span data-stu-id="f9bff-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9bff-165">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-165">Requirements</span></span>

|<span data-ttu-id="f9bff-166">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-166">Requirement</span></span>| <span data-ttu-id="f9bff-167">値</span><span class="sxs-lookup"><span data-stu-id="f9bff-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9bff-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f9bff-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9bff-169">1.0</span><span class="sxs-lookup"><span data-stu-id="f9bff-169">1.0</span></span>|
|[<span data-ttu-id="f9bff-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f9bff-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9bff-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9bff-171">ReadItem</span></span>|
|[<span data-ttu-id="f9bff-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f9bff-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9bff-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f9bff-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9bff-174">例</span><span class="sxs-lookup"><span data-stu-id="f9bff-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="f9bff-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="f9bff-175">timeZone :String</span></span>

<span data-ttu-id="f9bff-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f9bff-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="f9bff-177">型:</span><span class="sxs-lookup"><span data-stu-id="f9bff-177">Type:</span></span>

*   <span data-ttu-id="f9bff-178">String</span><span class="sxs-lookup"><span data-stu-id="f9bff-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9bff-179">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-179">Requirements</span></span>

|<span data-ttu-id="f9bff-180">要件</span><span class="sxs-lookup"><span data-stu-id="f9bff-180">Requirement</span></span>| <span data-ttu-id="f9bff-181">値</span><span class="sxs-lookup"><span data-stu-id="f9bff-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9bff-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f9bff-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9bff-183">1.0</span><span class="sxs-lookup"><span data-stu-id="f9bff-183">1.0</span></span>|
|[<span data-ttu-id="f9bff-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f9bff-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f9bff-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f9bff-185">ReadItem</span></span>|
|[<span data-ttu-id="f9bff-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f9bff-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9bff-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f9bff-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f9bff-188">例</span><span class="sxs-lookup"><span data-stu-id="f9bff-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```