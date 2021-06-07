---
title: アドインで代理アクセスシナリオOutlook有効にする
description: 委任アクセスについて簡単に説明し、アドインのサポートを構成する方法について説明します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 256c37087b10eaf9c8025e19a4990852f9550458
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783492"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="5a5f3-103">アドインで代理アクセスシナリオOutlook有効にする</span><span class="sxs-lookup"><span data-stu-id="5a5f3-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="5a5f3-104">メールボックス所有者は代理人アクセス機能を使用して、他のユーザーが自分のメールと予定表 [を管理できます](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="5a5f3-105">この記事では、JavaScript API がサポートする委任アクセス許可Office指定し、アドインで委任アクセス シナリオを有効にする方法Outlook説明します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5a5f3-106">委任アクセスは、Android および iOS Outlookでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="5a5f3-107">また、この機能は現在、Web 上のグループ共有[メールボックスOutlook使用](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)できません。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="5a5f3-108">この機能は、将来利用可能になる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="5a5f3-109">この機能のサポートは、要件セット 1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="5a5f3-110">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="5a5f3-111">代理人アクセスでサポートされているアクセス許可</span><span class="sxs-lookup"><span data-stu-id="5a5f3-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="5a5f3-112">次の表では、JavaScript API がサポートするOfficeアクセス許可について説明します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="5a5f3-113">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="5a5f3-113">Permission</span></span>|<span data-ttu-id="5a5f3-114">値</span><span class="sxs-lookup"><span data-stu-id="5a5f3-114">Value</span></span>|<span data-ttu-id="5a5f3-115">説明</span><span class="sxs-lookup"><span data-stu-id="5a5f3-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="5a5f3-116">読み取り</span><span class="sxs-lookup"><span data-stu-id="5a5f3-116">Read</span></span>|<span data-ttu-id="5a5f3-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="5a5f3-117">1 (000001)</span></span>|<span data-ttu-id="5a5f3-118">アイテムを読み取り可能。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-118">Can read items.</span></span>|
|<span data-ttu-id="5a5f3-119">書き込み</span><span class="sxs-lookup"><span data-stu-id="5a5f3-119">Write</span></span>|<span data-ttu-id="5a5f3-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="5a5f3-120">2 (000010)</span></span>|<span data-ttu-id="5a5f3-121">アイテムを作成できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-121">Can create items.</span></span>|
|<span data-ttu-id="5a5f3-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="5a5f3-122">DeleteOwn</span></span>|<span data-ttu-id="5a5f3-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="5a5f3-123">4 (000100)</span></span>|<span data-ttu-id="5a5f3-124">作成したアイテムのみを削除できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="5a5f3-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="5a5f3-125">DeleteAll</span></span>|<span data-ttu-id="5a5f3-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="5a5f3-126">8 (001000)</span></span>|<span data-ttu-id="5a5f3-127">任意のアイテムを削除できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-127">Can delete any items.</span></span>|
|<span data-ttu-id="5a5f3-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="5a5f3-128">EditOwn</span></span>|<span data-ttu-id="5a5f3-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="5a5f3-129">16 (010000)</span></span>|<span data-ttu-id="5a5f3-130">作成したアイテムのみを編集できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="5a5f3-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="5a5f3-131">EditAll</span></span>|<span data-ttu-id="5a5f3-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="5a5f3-132">32 (100000)</span></span>|<span data-ttu-id="5a5f3-133">任意のアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="5a5f3-134">現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人のアクセス許可は設定していない。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="5a5f3-135">[DelegatePermissions オブジェクト](/javascript/api/outlook/office.mailboxenums.delegatepermissions)は、代理人のアクセス許可を示すためにビットマスクを使用して実装されます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="5a5f3-136">ビットマスク内の各位置は特定のアクセス許可を表し、それが設定されている場合、代理人はそれぞれの `1` アクセス許可を持っています。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="5a5f3-137">たとえば、右側の 2 番目のビットがである場合、代理人 `1` は書き込みアクセス **許可を持** つ。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="5a5f3-138">特定のアクセス許可を確認する方法の例については、後の「[](#perform-an-operation-as-delegate)代理として操作を実行する」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="5a5f3-139">メールボックス クライアント間の同期</span><span class="sxs-lookup"><span data-stu-id="5a5f3-139">Sync across mailbox clients</span></span>

<span data-ttu-id="5a5f3-140">所有者のメールボックスに対する代理人の更新は、通常、すぐにメールボックス間で同期されます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="5a5f3-141">ただし、REST または Exchange Web サービス (EWS) 操作を使用してアイテムに拡張プロパティを設定した場合、このような変更は同期に数時間かかる可能性があります。このような遅延を回避するには[、CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトと関連する API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="5a5f3-142">詳細については、「Get [and](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) set metadata in the Outlook」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5a5f3-143">代理人シナリオでは、API によって現在提供されているトークンで EWS をoffice.jsできません。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="5a5f3-144">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="5a5f3-144">Configure the manifest</span></span>

<span data-ttu-id="5a5f3-145">アドインで委任アクセスシナリオを有効にするには、親要素の下のマニフェストで [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) 要素を `true` 設定する必要があります `DesktopFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="5a5f3-146">現時点では、他のフォーム ファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="5a5f3-147">代理人からの REST 呼び出しをサポートするには、マニフェストの [Permissions](../reference/manifest/permissions.md) ノードをに設定します `ReadWriteMailbox` 。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="5a5f3-148">次の例は、 `SupportsSharedFolders` マニフェストのセクションに `true` 設定された要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="5a5f3-149">デリゲートとして操作を実行する</span><span class="sxs-lookup"><span data-stu-id="5a5f3-149">Perform an operation as delegate</span></span>

<span data-ttu-id="5a5f3-150">[item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出すことによって、作成モードまたは読み取りモードでアイテムの共有プロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="5a5f3-151">これは、代理人のアクセス許可、所有者の電子メール アドレス、REST API の基本 URL、およびターゲット メールボックスを現在提供している [SharedProperties](/javascript/api/outlook/office.sharedproperties) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="5a5f3-152">次の例は、メッセージまたは予定の共有プロパティを取得し、代理人が **書** き込みアクセス許可を持つか確認し、REST 呼び出しを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="5a5f3-153">代理人として REST を使用して、アイテムまたはグループの投稿にOutlookメッセージのOutlook[を取得できます](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="5a5f3-154">共有アイテムと非共有アイテムの REST の呼び出しを処理する</span><span class="sxs-lookup"><span data-stu-id="5a5f3-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="5a5f3-155">アイテムに対して REST 操作を呼び出す場合は、アイテムが共有されるかどうかに関して、API を使用してアイテムが共有 `getSharedPropertiesAsync` されているかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="5a5f3-156">その後、適切なオブジェクトを使用して操作の REST URL を作成できます。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

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

## <a name="limitations"></a><span data-ttu-id="5a5f3-157">制限事項</span><span class="sxs-lookup"><span data-stu-id="5a5f3-157">Limitations</span></span>

<span data-ttu-id="5a5f3-158">アドインのシナリオに応じて、代理人の状況を処理する際に考慮する必要があるいくつかの制限があります。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="5a5f3-159">REST と EWS</span><span class="sxs-lookup"><span data-stu-id="5a5f3-159">REST and EWS</span></span>

<span data-ttu-id="5a5f3-160">アドインは REST を使用できますが、EWS は使用できません。また、所有者のメールボックスへの REST アクセスを有効にするには、アドインのアクセス許可 `ReadWriteMailbox` を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="5a5f3-161">メッセージ作成モード</span><span class="sxs-lookup"><span data-stu-id="5a5f3-161">Message Compose mode</span></span>

<span data-ttu-id="5a5f3-162">メッセージ作成モードでは[、getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_)は、以下の条件を満たしていない限り、Outlook または Windows でサポートされません。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="5a5f3-163">所有者は、代理人と少なくとも 1 つのメールボックス フォルダーを共有します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="5a5f3-164">代理人は、共有フォルダー内のメッセージを下書きします。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="5a5f3-165">例:</span><span class="sxs-lookup"><span data-stu-id="5a5f3-165">Examples:</span></span>

    - <span data-ttu-id="5a5f3-166">代理人は、共有フォルダー内の電子メールに返信または転送します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="5a5f3-167">代理人は下書きメッセージを保存し、それを自分の **下書き** フォルダーから共有フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="5a5f3-168">代理人は、共有フォルダーから下書きを開き、作成を続行します。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="5a5f3-169">メッセージが送信された後、通常は代理人の [送信されたアイテム] **フォルダーにあります** 。</span><span class="sxs-lookup"><span data-stu-id="5a5f3-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="5a5f3-170">関連項目</span><span class="sxs-lookup"><span data-stu-id="5a5f3-170">See also</span></span>

- [<span data-ttu-id="5a5f3-171">自分のメールと予定表の管理を他のユーザーに許可する</span><span class="sxs-lookup"><span data-stu-id="5a5f3-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="5a5f3-172">カレンダーの共有Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="5a5f3-172">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="5a5f3-173">マニフェスト要素を順序付けする方法</span><span class="sxs-lookup"><span data-stu-id="5a5f3-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="5a5f3-174">[マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="5a5f3-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="5a5f3-175">JavaScript ビット演算子</span><span class="sxs-lookup"><span data-stu-id="5a5f3-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)