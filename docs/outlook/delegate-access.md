---
title: Outlook アドインで代理人アクセスのシナリオを有効にする
description: 代理人アクセスについて簡単に説明し、アドインサポートを構成する方法について説明します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 598f931dbf3a4be8adf029838084ec0767bf6518
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234241"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="6cc9a-103">Outlook アドインで代理人アクセスのシナリオを有効にする</span><span class="sxs-lookup"><span data-stu-id="6cc9a-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="6cc9a-104">メールボックスの所有者は、代理人アクセス機能を使用して、他のユーザーが自分のメールと予定表 [を管理できます](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="6cc9a-105">この記事では、Office JavaScript API がサポートする委任アクセス許可を指定し、Outlook アドインで代理人アクセスシナリオを有効にする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6cc9a-106">Android および iOS 上の Outlook では、代理人アクセスは現在使用できません。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="6cc9a-107">また、この機能は、Outlook on the web の [グループ共有](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) メールボックスでは現在使用できません。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="6cc9a-108">この機能は、今後使用可能になる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="6cc9a-109">この機能のサポートは、要件セット 1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="6cc9a-110">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="6cc9a-111">代理人アクセスでサポートされているアクセス許可</span><span class="sxs-lookup"><span data-stu-id="6cc9a-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="6cc9a-112">次の表では、JavaScript API がサポートする委任Office説明します。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="6cc9a-113">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="6cc9a-113">Permission</span></span>|<span data-ttu-id="6cc9a-114">値</span><span class="sxs-lookup"><span data-stu-id="6cc9a-114">Value</span></span>|<span data-ttu-id="6cc9a-115">説明</span><span class="sxs-lookup"><span data-stu-id="6cc9a-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="6cc9a-116">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cc9a-116">Read</span></span>|<span data-ttu-id="6cc9a-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="6cc9a-117">1 (000001)</span></span>|<span data-ttu-id="6cc9a-118">アイテムを読み取り可能。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-118">Can read items.</span></span>|
|<span data-ttu-id="6cc9a-119">書き込み</span><span class="sxs-lookup"><span data-stu-id="6cc9a-119">Write</span></span>|<span data-ttu-id="6cc9a-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="6cc9a-120">2 (000010)</span></span>|<span data-ttu-id="6cc9a-121">アイテムを作成できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-121">Can create items.</span></span>|
|<span data-ttu-id="6cc9a-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="6cc9a-122">DeleteOwn</span></span>|<span data-ttu-id="6cc9a-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="6cc9a-123">4 (000100)</span></span>|<span data-ttu-id="6cc9a-124">作成したアイテムのみを削除できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="6cc9a-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="6cc9a-125">DeleteAll</span></span>|<span data-ttu-id="6cc9a-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="6cc9a-126">8 (001000)</span></span>|<span data-ttu-id="6cc9a-127">任意のアイテムを削除できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-127">Can delete any items.</span></span>|
|<span data-ttu-id="6cc9a-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="6cc9a-128">EditOwn</span></span>|<span data-ttu-id="6cc9a-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="6cc9a-129">16 (010000)</span></span>|<span data-ttu-id="6cc9a-130">作成したアイテムのみを編集できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="6cc9a-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="6cc9a-131">EditAll</span></span>|<span data-ttu-id="6cc9a-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="6cc9a-132">32 (100000)</span></span>|<span data-ttu-id="6cc9a-133">任意のアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="6cc9a-134">現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人のアクセス許可を設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="6cc9a-135">[DelegatePermissions オブジェクトは](/javascript/api/outlook/office.mailboxenums.delegatepermissions)、代理人のアクセス許可を示すためにビットマスクを使用して実装されます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="6cc9a-136">ビットマスク内の各位置は特定のアクセス許可を表し、設定されている場合、代理人はそれぞれのアクセス `1` 許可を持っています。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="6cc9a-137">たとえば、右側の 2 番目のビットが次の場合、代理人は書き込 `1` みアクセス **許可を持** つ。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="6cc9a-138">特定のアクセス許可を確認する方法の例については、この記事の[](#perform-an-operation-as-delegate)後半の「代理として操作を実行する」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="6cc9a-139">メールボックス クライアント間での同期</span><span class="sxs-lookup"><span data-stu-id="6cc9a-139">Sync across mailbox clients</span></span>

<span data-ttu-id="6cc9a-140">所有者のメールボックスに対する代理人の更新は、通常、メールボックス間で直ちに同期されます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="6cc9a-141">ただし、REST 操作または Exchange Web サービス (EWS) 操作を使用してアイテムに拡張プロパティを設定した場合、このような変更の同期には数時間かかる場合があります。このような遅延を避けるために [、CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトと関連する API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="6cc9a-142">詳細については、「Outlook アドイン [でメタデータ](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) を取得および設定する」の「カスタム プロパティ」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6cc9a-143">委任シナリオでは、現在 API によって提供されているトークンと EWS office.jsできません。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="6cc9a-144">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="6cc9a-144">Configure the manifest</span></span>

<span data-ttu-id="6cc9a-145">アドインで代理人アクセスのシナリオを有効にするには、親要素のマニフェストで [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) 要素を設定 `true` する必要があります `DesktopFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="6cc9a-146">現時点では、他のフォーム ファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="6cc9a-147">代理人からの REST 呼び出しをサポートするには、マニフェスト [のアクセス](../reference/manifest/permissions.md) 許可ノードを次に設定します `ReadWriteMailbox` 。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="6cc9a-148">次の例は、 `SupportsSharedFolders` マニフェストのセクション `true` に設定されている要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="6cc9a-149">デリゲートとして操作を実行する</span><span class="sxs-lookup"><span data-stu-id="6cc9a-149">Perform an operation as delegate</span></span>

<span data-ttu-id="6cc9a-150">[item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出すことによって、新規作成モードまたは読み取りモードでアイテムの共有プロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="6cc9a-151">これにより、代理人のアクセス許可、所有者の電子メール アドレス、REST API のベース URL、およびターゲット メールボックスを現在提供している [SharedProperties](/javascript/api/outlook/office.sharedproperties) オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="6cc9a-152">次の例は、メッセージまたは予定の共有プロパティを取得し、代理人が **書** き込みアクセス許可を持って、REST 呼び出しを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="6cc9a-153">代理人は、REST を使用して、Outlook アイテムまたはグループの投稿に添付された Outlook メッセージのコンテンツ [を取得できます](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="6cc9a-154">共有アイテムと非共有アイテムに対する REST の呼び出しの処理</span><span class="sxs-lookup"><span data-stu-id="6cc9a-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="6cc9a-155">アイテムが共有されるかどうかに関して、アイテムに対して REST 操作を呼び出す場合は、API を使用してアイテムが共有 `getSharedPropertiesAsync` されているかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="6cc9a-156">その後、適切なオブジェクトを使用して操作の REST URL を作成できます。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="6cc9a-157">制限事項</span><span class="sxs-lookup"><span data-stu-id="6cc9a-157">Limitations</span></span>

<span data-ttu-id="6cc9a-158">アドインのシナリオに応じて、代理人の状況を処理する際に考慮する必要がある制限があります。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="6cc9a-159">REST と EWS</span><span class="sxs-lookup"><span data-stu-id="6cc9a-159">REST and EWS</span></span>

<span data-ttu-id="6cc9a-160">アドインは REST を使用できますが、EWS は使用できません。また、所有者のメールボックスへの REST アクセスを有効にするには、アドインのアクセス許可を設定 `ReadWriteMailbox` する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="6cc9a-161">メッセージ作成モード</span><span class="sxs-lookup"><span data-stu-id="6cc9a-161">Message Compose mode</span></span>

<span data-ttu-id="6cc9a-162">メッセージ作成モードでは、次の条件が満たされない限り、Outlook on the web または Windows では [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) はサポートされません。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="6cc9a-163">所有者は、代理人と少なくとも 1 つのメールボックス フォルダーを共有します。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="6cc9a-164">代理人は、共有フォルダー内のメッセージを下書きします。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="6cc9a-165">例:</span><span class="sxs-lookup"><span data-stu-id="6cc9a-165">Examples:</span></span>

    - <span data-ttu-id="6cc9a-166">代理人は、共有フォルダー内の電子メールに返信または転送します。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="6cc9a-167">代理人は下書きメッセージを保存し、それを自分の **下書き** フォルダーから共有フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="6cc9a-168">代理人は、共有フォルダーから下書きを開き、作成を続行します。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="6cc9a-169">メッセージが送信された後は、通常、代理人の [送信されたアイテム] **フォルダーにあります** 。</span><span class="sxs-lookup"><span data-stu-id="6cc9a-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="6cc9a-170">関連項目</span><span class="sxs-lookup"><span data-stu-id="6cc9a-170">See also</span></span>

- [<span data-ttu-id="6cc9a-171">自分のメールと予定表の管理を他のユーザーに許可する</span><span class="sxs-lookup"><span data-stu-id="6cc9a-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="6cc9a-172">Microsoft 365 での予定表の共有</span><span class="sxs-lookup"><span data-stu-id="6cc9a-172">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="6cc9a-173">マニフェスト要素を順序付けする方法</span><span class="sxs-lookup"><span data-stu-id="6cc9a-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="6cc9a-174">[マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="6cc9a-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="6cc9a-175">JavaScript のビット演算子</span><span class="sxs-lookup"><span data-stu-id="6cc9a-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)