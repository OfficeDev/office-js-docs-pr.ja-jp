---
title: Outlook アドインで代理人アクセスのシナリオを有効にする
description: 代理人アクセスについて簡単に説明し、アドインサポートを構成する方法について説明します。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 68b9e09afbe2bcd5cfc302d6714b1c22fd945047
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608951"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="5261b-103">Outlook アドインで代理人アクセスのシナリオを有効にする</span><span class="sxs-lookup"><span data-stu-id="5261b-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="5261b-104">メールボックスの所有者は代理人アクセス機能を使用して、[他のユーザーが自分のメールと予定表を管理できるよう](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)にすることができます。</span><span class="sxs-lookup"><span data-stu-id="5261b-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="5261b-105">この記事では、Office JavaScript API でサポートされている代理人アクセス許可を指定し、Outlook アドインで代理人アクセスのシナリオを有効にする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="5261b-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5261b-106">現在、Outlook on Mac、Android、iOS では、代理人アクセスは利用できません。</span><span class="sxs-lookup"><span data-stu-id="5261b-106">Delegate access is not currently available in Outlook on Mac, Android, and iOS.</span></span> <span data-ttu-id="5261b-107">この機能は、今後利用可能になる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="5261b-107">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="5261b-108">この機能のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="5261b-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="5261b-109">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5261b-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="5261b-110">代理人アクセスに対してサポートされるアクセス許可</span><span class="sxs-lookup"><span data-stu-id="5261b-110">Supported permissions for delegate access</span></span>

<span data-ttu-id="5261b-111">次の表では、Office JavaScript API でサポートされている代理人アクセス許可について説明します。</span><span class="sxs-lookup"><span data-stu-id="5261b-111">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="5261b-112">Permission</span><span class="sxs-lookup"><span data-stu-id="5261b-112">Permission</span></span>|<span data-ttu-id="5261b-113">値</span><span class="sxs-lookup"><span data-stu-id="5261b-113">Value</span></span>|<span data-ttu-id="5261b-114">説明</span><span class="sxs-lookup"><span data-stu-id="5261b-114">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="5261b-115">Read</span><span class="sxs-lookup"><span data-stu-id="5261b-115">Read</span></span>|<span data-ttu-id="5261b-116">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="5261b-116">1 (000001)</span></span>|<span data-ttu-id="5261b-117">アイテムを読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="5261b-117">Can read items.</span></span>|
|<span data-ttu-id="5261b-118">書き込み</span><span class="sxs-lookup"><span data-stu-id="5261b-118">Write</span></span>|<span data-ttu-id="5261b-119">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="5261b-119">2 (000010)</span></span>|<span data-ttu-id="5261b-120">アイテムを作成できます。</span><span class="sxs-lookup"><span data-stu-id="5261b-120">Can create items.</span></span>|
|<span data-ttu-id="5261b-121">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="5261b-121">DeleteOwn</span></span>|<span data-ttu-id="5261b-122">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="5261b-122">4 (000100)</span></span>|<span data-ttu-id="5261b-123">は、自分で作成したアイテムのみを削除できます。</span><span class="sxs-lookup"><span data-stu-id="5261b-123">Can delete only the items they created.</span></span>|
|<span data-ttu-id="5261b-124">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="5261b-124">DeleteAll</span></span>|<span data-ttu-id="5261b-125">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="5261b-125">8 (001000)</span></span>|<span data-ttu-id="5261b-126">任意のアイテムを削除できます。</span><span class="sxs-lookup"><span data-stu-id="5261b-126">Can delete any items.</span></span>|
|<span data-ttu-id="5261b-127">EditOwn</span><span class="sxs-lookup"><span data-stu-id="5261b-127">EditOwn</span></span>|<span data-ttu-id="5261b-128">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="5261b-128">16 (010000)</span></span>|<span data-ttu-id="5261b-129">は、自分で作成したアイテムのみを編集できます。</span><span class="sxs-lookup"><span data-stu-id="5261b-129">Can edit only the items they created.</span></span>|
|<span data-ttu-id="5261b-130">EditAll</span><span class="sxs-lookup"><span data-stu-id="5261b-130">EditAll</span></span>|<span data-ttu-id="5261b-131">32 (10万)</span><span class="sxs-lookup"><span data-stu-id="5261b-131">32 (100000)</span></span>|<span data-ttu-id="5261b-132">任意のアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="5261b-132">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="5261b-133">現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人アクセス許可は設定しません。</span><span class="sxs-lookup"><span data-stu-id="5261b-133">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="5261b-134">[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)オブジェクトは、デリゲートのアクセス許可を示すために、ビットマスクを使用して実装されます。</span><span class="sxs-lookup"><span data-stu-id="5261b-134">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="5261b-135">ビットマスク内の各位置は特定のアクセス許可を表し、それが設定されている場合は `1` 代理人にそれぞれのアクセス許可が付与されます。</span><span class="sxs-lookup"><span data-stu-id="5261b-135">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="5261b-136">たとえば、右側の2番目のビットがの場合、 `1` デリゲートには**書き込み**アクセス許可があります。</span><span class="sxs-lookup"><span data-stu-id="5261b-136">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="5261b-137">この記事で後述する「[代理人として操作を実行](#perform-an-operation-as-delegate)する」の特定のアクセス許可を確認する方法の例を確認できます。</span><span class="sxs-lookup"><span data-stu-id="5261b-137">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="5261b-138">メールボックスクライアント間での同期</span><span class="sxs-lookup"><span data-stu-id="5261b-138">Sync across mailbox clients</span></span>

<span data-ttu-id="5261b-139">通常、所有者のメールボックスに対する代理人の更新は、メールボックス間で即時に同期されます。</span><span class="sxs-lookup"><span data-stu-id="5261b-139">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="5261b-140">ただし、アドインが REST または EWS 操作を使用してアイテムの拡張プロパティを設定する場合、そのような変更は同期に数時間かかることがあります。このような遅延を避けるには、 [CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトと関連する api を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="5261b-140">However, if the add-in uses REST or EWS operations to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="5261b-141">詳細については、記事「Outlook アドインでメタデータを取得および設定する」の「[カスタムプロパティ」セクション](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5261b-141">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="5261b-142">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="5261b-142">Configure the manifest</span></span>

<span data-ttu-id="5261b-143">アドインで代理人アクセスのシナリオを有効にするには、親要素のマニフェスト内の[Supportssharedfolders](../reference/manifest/supportssharedfolders.md)要素をに設定する必要があり `true` `DesktopFormFactor` ます。</span><span class="sxs-lookup"><span data-stu-id="5261b-143">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="5261b-144">現在、他のフォームファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5261b-144">At present, other form factors are not supported.</span></span>

<span data-ttu-id="5261b-145">次の例は、 `SupportsSharedFolders` マニフェストのセクション内に設定された要素を示して `true` います。</span><span class="sxs-lookup"><span data-stu-id="5261b-145">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="5261b-146">代理人として操作を実行する</span><span class="sxs-lookup"><span data-stu-id="5261b-146">Perform an operation as delegate</span></span>

<span data-ttu-id="5261b-147">アイテムの共有プロパティは、新規作成または閲覧モードで取得できます。そのためには、 [getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="5261b-147">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="5261b-148">これにより、現在、代理人のアクセス許可、所有者の電子メールアドレス、REST API のベース URL、ターゲットメールボックスを提供する[Sharedproperties](/javascript/api/outlook/office.sharedproperties)オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="5261b-148">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="5261b-149">次の例は、メッセージまたは予定の共有プロパティを取得する方法、代理人が**書き込み**アクセス許可を持っているかどうかを確認する方法、および REST 呼び出しを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5261b-149">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="5261b-150">関連項目</span><span class="sxs-lookup"><span data-stu-id="5261b-150">See also</span></span>

- [<span data-ttu-id="5261b-151">自分のメールと予定表の管理を他のユーザーに許可する</span><span class="sxs-lookup"><span data-stu-id="5261b-151">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="5261b-152">Office365 での予定表の共有</span><span class="sxs-lookup"><span data-stu-id="5261b-152">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="5261b-153">マニフェスト要素の注文方法</span><span class="sxs-lookup"><span data-stu-id="5261b-153">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="5261b-154">[マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="5261b-154">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="5261b-155">JavaScript ビット演算子</span><span class="sxs-lookup"><span data-stu-id="5261b-155">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)