---
title: Outlook アドインで代理人アクセスのシナリオを有効にする
description: 代理人アクセスについて簡単に説明し、アドインサポートを構成する方法について説明します。
ms.date: 06/30/2020
localization_priority: Normal
ms.openlocfilehash: a5b4581783ca65bfe858dcf6638287418a3dcfe2
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006417"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="35c23-103">Outlook アドインで代理人アクセスのシナリオを有効にする</span><span class="sxs-lookup"><span data-stu-id="35c23-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="35c23-104">メールボックスの所有者は代理人アクセス機能を使用して、[他のユーザーが自分のメールと予定表を管理できるよう](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)にすることができます。</span><span class="sxs-lookup"><span data-stu-id="35c23-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="35c23-105">この記事では、Office JavaScript API でサポートされている代理人アクセス許可を指定し、Outlook アドインで代理人アクセスのシナリオを有効にする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="35c23-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="35c23-106">現在、Outlook on Mac、Android、iOS では、代理人アクセスは利用できません。</span><span class="sxs-lookup"><span data-stu-id="35c23-106">Delegate access is not currently available in Outlook on Mac, Android, and iOS.</span></span> <span data-ttu-id="35c23-107">この機能は、今後利用可能になる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="35c23-107">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="35c23-108">この機能のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="35c23-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="35c23-109">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35c23-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="35c23-110">代理人アクセスに対してサポートされるアクセス許可</span><span class="sxs-lookup"><span data-stu-id="35c23-110">Supported permissions for delegate access</span></span>

<span data-ttu-id="35c23-111">次の表では、Office JavaScript API でサポートされている代理人アクセス許可について説明します。</span><span class="sxs-lookup"><span data-stu-id="35c23-111">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="35c23-112">Permission</span><span class="sxs-lookup"><span data-stu-id="35c23-112">Permission</span></span>|<span data-ttu-id="35c23-113">値</span><span class="sxs-lookup"><span data-stu-id="35c23-113">Value</span></span>|<span data-ttu-id="35c23-114">説明</span><span class="sxs-lookup"><span data-stu-id="35c23-114">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="35c23-115">Read</span><span class="sxs-lookup"><span data-stu-id="35c23-115">Read</span></span>|<span data-ttu-id="35c23-116">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="35c23-116">1 (000001)</span></span>|<span data-ttu-id="35c23-117">アイテムを読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="35c23-117">Can read items.</span></span>|
|<span data-ttu-id="35c23-118">書き込み</span><span class="sxs-lookup"><span data-stu-id="35c23-118">Write</span></span>|<span data-ttu-id="35c23-119">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="35c23-119">2 (000010)</span></span>|<span data-ttu-id="35c23-120">アイテムを作成できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-120">Can create items.</span></span>|
|<span data-ttu-id="35c23-121">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="35c23-121">DeleteOwn</span></span>|<span data-ttu-id="35c23-122">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="35c23-122">4 (000100)</span></span>|<span data-ttu-id="35c23-123">は、自分で作成したアイテムのみを削除できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-123">Can delete only the items they created.</span></span>|
|<span data-ttu-id="35c23-124">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="35c23-124">DeleteAll</span></span>|<span data-ttu-id="35c23-125">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="35c23-125">8 (001000)</span></span>|<span data-ttu-id="35c23-126">任意のアイテムを削除できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-126">Can delete any items.</span></span>|
|<span data-ttu-id="35c23-127">EditOwn</span><span class="sxs-lookup"><span data-stu-id="35c23-127">EditOwn</span></span>|<span data-ttu-id="35c23-128">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="35c23-128">16 (010000)</span></span>|<span data-ttu-id="35c23-129">は、自分で作成したアイテムのみを編集できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-129">Can edit only the items they created.</span></span>|
|<span data-ttu-id="35c23-130">EditAll</span><span class="sxs-lookup"><span data-stu-id="35c23-130">EditAll</span></span>|<span data-ttu-id="35c23-131">32 (10万)</span><span class="sxs-lookup"><span data-stu-id="35c23-131">32 (100000)</span></span>|<span data-ttu-id="35c23-132">任意のアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-132">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="35c23-133">現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人アクセス許可は設定しません。</span><span class="sxs-lookup"><span data-stu-id="35c23-133">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="35c23-134">[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)オブジェクトは、デリゲートのアクセス許可を示すために、ビットマスクを使用して実装されます。</span><span class="sxs-lookup"><span data-stu-id="35c23-134">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="35c23-135">ビットマスク内の各位置は特定のアクセス許可を表し、それが設定されている場合は `1` 代理人にそれぞれのアクセス許可が付与されます。</span><span class="sxs-lookup"><span data-stu-id="35c23-135">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="35c23-136">たとえば、右側の2番目のビットがの場合、 `1` デリゲートには**書き込み**アクセス許可があります。</span><span class="sxs-lookup"><span data-stu-id="35c23-136">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="35c23-137">この記事で後述する「[代理人として操作を実行](#perform-an-operation-as-delegate)する」の特定のアクセス許可を確認する方法の例を確認できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-137">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="35c23-138">メールボックスクライアント間での同期</span><span class="sxs-lookup"><span data-stu-id="35c23-138">Sync across mailbox clients</span></span>

<span data-ttu-id="35c23-139">通常、所有者のメールボックスに対する代理人の更新は、メールボックス間で即時に同期されます。</span><span class="sxs-lookup"><span data-stu-id="35c23-139">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="35c23-140">ただし、REST または Exchange Web サービス (EWS) の操作を使用してアイテムの拡張プロパティを設定した場合は、そのような変更を同期するのに数時間かかることがあります。このような遅延を避けるには、 [CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトと関連する api を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="35c23-140">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="35c23-141">詳細については、記事「Outlook アドインでメタデータを取得および設定する」の「[カスタムプロパティ」セクション](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="35c23-141">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="35c23-142">代理人のシナリオでは、office.js API によって現在提供されているトークンと共に EWS を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="35c23-142">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="35c23-143">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="35c23-143">Configure the manifest</span></span>

<span data-ttu-id="35c23-144">アドインで代理人アクセスのシナリオを有効にするには、親要素のマニフェスト内の[Supportssharedfolders](../reference/manifest/supportssharedfolders.md)要素をに設定する必要があり `true` `DesktopFormFactor` ます。</span><span class="sxs-lookup"><span data-stu-id="35c23-144">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="35c23-145">現在、他のフォームファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="35c23-145">At present, other form factors are not supported.</span></span>

<span data-ttu-id="35c23-146">代理人からの REST 呼び出しをサポートするには、マニフェストの [[アクセス許可](../reference/manifest/permissions.md)] ノードをに設定 `ReadWriteMailbox` します。</span><span class="sxs-lookup"><span data-stu-id="35c23-146">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="35c23-147">次の例は、 `SupportsSharedFolders` マニフェストのセクション内に設定された要素を示して `true` います。</span><span class="sxs-lookup"><span data-stu-id="35c23-147">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="35c23-148">代理人として操作を実行する</span><span class="sxs-lookup"><span data-stu-id="35c23-148">Perform an operation as delegate</span></span>

<span data-ttu-id="35c23-149">アイテムの共有プロパティは、新規作成または閲覧モードで取得できます。そのためには、 [getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="35c23-149">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="35c23-150">これにより、現在、代理人のアクセス許可、所有者の電子メールアドレス、REST API のベース URL、ターゲットメールボックスを提供する[Sharedproperties](/javascript/api/outlook/office.sharedproperties)オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="35c23-150">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="35c23-151">代理人のシナリオでは、アドインでは、EWS ではなく REST を使用でき、アドインのアクセス許可をに設定して、 `ReadWriteMailbox` 所有者のメールボックスへの rest アクセスを有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="35c23-151">In a delegate scenario, your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

<span data-ttu-id="35c23-152">次の例は、メッセージまたは予定の共有プロパティを取得する方法、代理人が**書き込み**アクセス許可を持っているかどうかを確認する方法、および REST 呼び出しを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="35c23-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="35c23-153">代理人は、REST を使用して、 [outlook アイテムまたはグループ投稿に添付されている outlook メッセージのコンテンツを取得](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)できます。</span><span class="sxs-lookup"><span data-stu-id="35c23-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="see-also"></a><span data-ttu-id="35c23-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="35c23-154">See also</span></span>

- [<span data-ttu-id="35c23-155">自分のメールと予定表の管理を他のユーザーに許可する</span><span class="sxs-lookup"><span data-stu-id="35c23-155">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="35c23-156">Office365 での予定表の共有</span><span class="sxs-lookup"><span data-stu-id="35c23-156">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="35c23-157">マニフェスト要素の注文方法</span><span class="sxs-lookup"><span data-stu-id="35c23-157">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="35c23-158">[マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="35c23-158">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="35c23-159">JavaScript ビット演算子</span><span class="sxs-lookup"><span data-stu-id="35c23-159">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)