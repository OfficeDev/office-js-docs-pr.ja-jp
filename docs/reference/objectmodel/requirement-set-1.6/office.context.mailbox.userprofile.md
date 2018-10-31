
# <a name="userprofile"></a><span data-ttu-id="d6ff5-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="d6ff5-101">userProfile</span></span>

### <span data-ttu-id="d6ff5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d6ff5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d6ff5-104">要件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-104">Requirements</span></span>

|<span data-ttu-id="d6ff5-105">必要条件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-105">Requirement</span></span>| <span data-ttu-id="d6ff5-106">値</span><span class="sxs-lookup"><span data-stu-id="d6ff5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d6ff5-107">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d6ff5-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d6ff5-108">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d6ff5-108">1.0</span></span>|
|[<span data-ttu-id="d6ff5-109">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d6ff5-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d6ff5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d6ff5-110">ReadItem</span></span>|
|[<span data-ttu-id="d6ff5-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d6ff5-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d6ff5-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d6ff5-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d6ff5-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d6ff5-113">Members and methods</span></span>

| <span data-ttu-id="d6ff5-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="d6ff5-114">Member</span></span> | <span data-ttu-id="d6ff5-115">種類</span><span class="sxs-lookup"><span data-stu-id="d6ff5-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="d6ff5-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="d6ff5-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="d6ff5-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="d6ff5-117">Member</span></span> |
| [<span data-ttu-id="d6ff5-118">displayName</span><span class="sxs-lookup"><span data-stu-id="d6ff5-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="d6ff5-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="d6ff5-119">Member</span></span> |
| [<span data-ttu-id="d6ff5-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d6ff5-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="d6ff5-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="d6ff5-121">Member</span></span> |
| [<span data-ttu-id="d6ff5-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="d6ff5-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="d6ff5-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="d6ff5-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d6ff5-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="d6ff5-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="d6ff5-125">accountType: 文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="d6ff5-126">このメンバーは、現在 Outlook 2016 for Mac またはそれ以降でのみサポートされています (ビルド 16.9.1212 またはそれ以降)。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="d6ff5-p102">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。使用可能な値は、次の表に表示されます。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-p102">Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.</span></span>

| <span data-ttu-id="d6ff5-129">値</span><span class="sxs-lookup"><span data-stu-id="d6ff5-129">Value</span></span> | <span data-ttu-id="d6ff5-130">説明</span><span class="sxs-lookup"><span data-stu-id="d6ff5-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="d6ff5-131">メールボックスは、オンプレミスの Exchange Server にあります。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="d6ff5-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="d6ff5-133">メールボックスは、Office 365 の職場や学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="d6ff5-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="d6ff5-135">種類:</span><span class="sxs-lookup"><span data-stu-id="d6ff5-135">Type:</span></span>

*   <span data-ttu-id="d6ff5-136">文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d6ff5-137">要件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-137">Requirements</span></span>

|<span data-ttu-id="d6ff5-138">必要条件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-138">Requirement</span></span>| <span data-ttu-id="d6ff5-139">値</span><span class="sxs-lookup"><span data-stu-id="d6ff5-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="d6ff5-140">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d6ff5-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d6ff5-141">1.6</span><span class="sxs-lookup"><span data-stu-id="d6ff5-141">-16</span></span> |
|[<span data-ttu-id="d6ff5-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d6ff5-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d6ff5-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d6ff5-143">ReadItem</span></span>|
|[<span data-ttu-id="d6ff5-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d6ff5-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d6ff5-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d6ff5-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d6ff5-146">例</span><span class="sxs-lookup"><span data-stu-id="d6ff5-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="d6ff5-147">displayName: 文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-147">displayName :String</span></span>

<span data-ttu-id="d6ff5-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d6ff5-149">種類:</span><span class="sxs-lookup"><span data-stu-id="d6ff5-149">Type:</span></span>

*   <span data-ttu-id="d6ff5-150">文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d6ff5-151">要件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-151">Requirements</span></span>

|<span data-ttu-id="d6ff5-152">必要条件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-152">Requirement</span></span>| <span data-ttu-id="d6ff5-153">値</span><span class="sxs-lookup"><span data-stu-id="d6ff5-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="d6ff5-154">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d6ff5-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d6ff5-155">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d6ff5-155">1.0</span></span>|
|[<span data-ttu-id="d6ff5-156">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d6ff5-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d6ff5-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d6ff5-157">ReadItem</span></span>|
|[<span data-ttu-id="d6ff5-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d6ff5-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d6ff5-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d6ff5-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d6ff5-160">例</span><span class="sxs-lookup"><span data-stu-id="d6ff5-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d6ff5-161">emailAddress : 文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-161">emailAddress :String</span></span>

<span data-ttu-id="d6ff5-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d6ff5-163">種類:</span><span class="sxs-lookup"><span data-stu-id="d6ff5-163">Type:</span></span>

*   <span data-ttu-id="d6ff5-164">文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d6ff5-165">要件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-165">Requirements</span></span>

|<span data-ttu-id="d6ff5-166">必要条件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-166">Requirement</span></span>| <span data-ttu-id="d6ff5-167">値</span><span class="sxs-lookup"><span data-stu-id="d6ff5-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d6ff5-168">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d6ff5-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d6ff5-169">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d6ff5-169">1.0</span></span>|
|[<span data-ttu-id="d6ff5-170">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d6ff5-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d6ff5-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d6ff5-171">ReadItem</span></span>|
|[<span data-ttu-id="d6ff5-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d6ff5-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d6ff5-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d6ff5-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d6ff5-174">例</span><span class="sxs-lookup"><span data-stu-id="d6ff5-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d6ff5-175">タイム ゾーン : 文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-175">timeZone :String</span></span>

<span data-ttu-id="d6ff5-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d6ff5-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d6ff5-177">種類:</span><span class="sxs-lookup"><span data-stu-id="d6ff5-177">Type:</span></span>

*   <span data-ttu-id="d6ff5-178">文字列</span><span class="sxs-lookup"><span data-stu-id="d6ff5-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d6ff5-179">要件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-179">Requirements</span></span>

|<span data-ttu-id="d6ff5-180">必要条件</span><span class="sxs-lookup"><span data-stu-id="d6ff5-180">Requirement</span></span>| <span data-ttu-id="d6ff5-181">値</span><span class="sxs-lookup"><span data-stu-id="d6ff5-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d6ff5-182">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="d6ff5-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d6ff5-183">1.0以降</span><span class="sxs-lookup"><span data-stu-id="d6ff5-183">1.0</span></span>|
|[<span data-ttu-id="d6ff5-184">アクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d6ff5-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d6ff5-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d6ff5-185">ReadItem</span></span>|
|[<span data-ttu-id="d6ff5-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d6ff5-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d6ff5-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d6ff5-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d6ff5-188">例</span><span class="sxs-lookup"><span data-stu-id="d6ff5-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```