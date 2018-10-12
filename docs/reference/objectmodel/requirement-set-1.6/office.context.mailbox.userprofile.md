
# <a name="userprofile"></a><span data-ttu-id="c11d0-101">ユーザー プロファイル</span><span class="sxs-lookup"><span data-stu-id="c11d0-101">userProfile</span></span>

### <span data-ttu-id="c11d0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md)。ユーザー プロファイル</span><span class="sxs-lookup"><span data-stu-id="c11d0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c11d0-104">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-104">Requirements</span></span>

|<span data-ttu-id="c11d0-105">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-105">Requirement</span></span>| <span data-ttu-id="c11d0-106">値</span><span class="sxs-lookup"><span data-stu-id="c11d0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c11d0-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c11d0-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c11d0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c11d0-108">1.0</span></span>|
|[<span data-ttu-id="c11d0-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c11d0-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c11d0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c11d0-110">ReadItem</span></span>|
|[<span data-ttu-id="c11d0-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c11d0-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c11d0-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c11d0-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c11d0-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c11d0-113">Members and methods</span></span>

| <span data-ttu-id="c11d0-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="c11d0-114">Member</span></span> | <span data-ttu-id="c11d0-115">型</span><span class="sxs-lookup"><span data-stu-id="c11d0-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="c11d0-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="c11d0-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="c11d0-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="c11d0-117">Member</span></span> |
| [<span data-ttu-id="c11d0-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c11d0-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c11d0-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="c11d0-119">Member</span></span> |
| [<span data-ttu-id="c11d0-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c11d0-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c11d0-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="c11d0-121">Member</span></span> |
| [<span data-ttu-id="c11d0-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c11d0-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c11d0-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="c11d0-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c11d0-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="c11d0-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c11d0-125">accountType: 文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c11d0-126">このメンバーは、現在 Outlook 2016 for Mac またはそれ以降でのみサポートされています (ビルド 16.9.1212 またはそれ以降)。</span><span class="sxs-lookup"><span data-stu-id="c11d0-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="c11d0-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c11d0-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c11d0-128">使用可能な値は、次の表に表示されます。</span><span class="sxs-lookup"><span data-stu-id="c11d0-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="c11d0-129">値</span><span class="sxs-lookup"><span data-stu-id="c11d0-129">Value</span></span> | <span data-ttu-id="c11d0-130">説明</span><span class="sxs-lookup"><span data-stu-id="c11d0-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c11d0-131">メールボックスは、オンプレミスの Exchange Server にあります。</span><span class="sxs-lookup"><span data-stu-id="c11d0-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c11d0-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="c11d0-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c11d0-133">メールボックスは、Office 365 の職場や学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="c11d0-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c11d0-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="c11d0-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c11d0-135">型:</span><span class="sxs-lookup"><span data-stu-id="c11d0-135">Type:</span></span>

*   <span data-ttu-id="c11d0-136">文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c11d0-137">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-137">Requirements</span></span>

|<span data-ttu-id="c11d0-138">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-138">Requirement</span></span>| <span data-ttu-id="c11d0-139">値</span><span class="sxs-lookup"><span data-stu-id="c11d0-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c11d0-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c11d0-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c11d0-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c11d0-141">-16</span></span> |
|[<span data-ttu-id="c11d0-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c11d0-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c11d0-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c11d0-143">ReadItem</span></span>|
|[<span data-ttu-id="c11d0-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c11d0-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c11d0-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c11d0-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c11d0-146">例</span><span class="sxs-lookup"><span data-stu-id="c11d0-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="c11d0-147">displayName: 文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-147">displayName :String</span></span>

<span data-ttu-id="c11d0-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c11d0-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c11d0-149">型:</span><span class="sxs-lookup"><span data-stu-id="c11d0-149">Type:</span></span>

*   <span data-ttu-id="c11d0-150">文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c11d0-151">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-151">Requirements</span></span>

|<span data-ttu-id="c11d0-152">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-152">Requirement</span></span>| <span data-ttu-id="c11d0-153">値</span><span class="sxs-lookup"><span data-stu-id="c11d0-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c11d0-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c11d0-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c11d0-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c11d0-155">1.0</span></span>|
|[<span data-ttu-id="c11d0-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c11d0-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c11d0-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c11d0-157">ReadItem</span></span>|
|[<span data-ttu-id="c11d0-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c11d0-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c11d0-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c11d0-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c11d0-160">例</span><span class="sxs-lookup"><span data-stu-id="c11d0-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c11d0-161">emailAddress: 文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-161">emailAddress :String</span></span>

<span data-ttu-id="c11d0-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c11d0-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c11d0-163">型:</span><span class="sxs-lookup"><span data-stu-id="c11d0-163">Type:</span></span>

*   <span data-ttu-id="c11d0-164">文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c11d0-165">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-165">Requirements</span></span>

|<span data-ttu-id="c11d0-166">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-166">Requirement</span></span>| <span data-ttu-id="c11d0-167">値</span><span class="sxs-lookup"><span data-stu-id="c11d0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c11d0-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c11d0-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c11d0-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c11d0-169">1.0</span></span>|
|[<span data-ttu-id="c11d0-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c11d0-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c11d0-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c11d0-171">ReadItem</span></span>|
|[<span data-ttu-id="c11d0-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c11d0-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c11d0-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c11d0-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c11d0-174">例</span><span class="sxs-lookup"><span data-stu-id="c11d0-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c11d0-175">タイム ゾーン: 文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-175">timeZone :String</span></span>

<span data-ttu-id="c11d0-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c11d0-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c11d0-177">型:</span><span class="sxs-lookup"><span data-stu-id="c11d0-177">Type:</span></span>

*   <span data-ttu-id="c11d0-178">文字列</span><span class="sxs-lookup"><span data-stu-id="c11d0-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c11d0-179">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-179">Requirements</span></span>

|<span data-ttu-id="c11d0-180">要件</span><span class="sxs-lookup"><span data-stu-id="c11d0-180">Requirement</span></span>| <span data-ttu-id="c11d0-181">値</span><span class="sxs-lookup"><span data-stu-id="c11d0-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c11d0-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c11d0-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c11d0-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c11d0-183">1.0</span></span>|
|[<span data-ttu-id="c11d0-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c11d0-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c11d0-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c11d0-185">ReadItem</span></span>|
|[<span data-ttu-id="c11d0-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c11d0-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c11d0-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c11d0-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c11d0-188">例</span><span class="sxs-lookup"><span data-stu-id="c11d0-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```