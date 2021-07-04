---
title: 共有フォルダーと共有メールボックスのシナリオを、Outlookアドインで有効にする
description: 共有フォルダー (a.k.a) のアドイン サポートを構成する方法について説明します。 アクセスを委任する) と共有メールボックスを使用します。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 70578f2c78a9dd88efc9ba70d5599a13e121df53
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290713"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="6017e-104">共有フォルダーと共有メールボックスのシナリオを、Outlookアドインで有効にする</span><span class="sxs-lookup"><span data-stu-id="6017e-104">Enable shared folders and shared mailbox scenarios in an Outlook add-in</span></span>

<span data-ttu-id="6017e-105">この記事では、Outlook アドインで共有フォルダー (代理人アクセスとも呼ばれる) と共有メールボックス (プレビュー[中)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)のシナリオ (Office JavaScript API がサポートするアクセス許可を含む) を有効にする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6017e-105">This article describes how to enable shared folders (also known as delegate access) and shared mailbox (now in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)) scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6017e-106">この機能のサポートは、要件セット [1.8 で導入されました](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)。</span><span class="sxs-lookup"><span data-stu-id="6017e-106">Support for this feature was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="6017e-107">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6017e-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-setups"></a><span data-ttu-id="6017e-108">サポートされているセットアップ</span><span class="sxs-lookup"><span data-stu-id="6017e-108">Supported setups</span></span>

<span data-ttu-id="6017e-109">次のセクションでは、共有メールボックス (プレビュー中) と共有フォルダーでサポートされる構成について説明します。</span><span class="sxs-lookup"><span data-stu-id="6017e-109">The following sections describe supported configurations for shared mailboxes (now in preview) and shared folders.</span></span> <span data-ttu-id="6017e-110">機能 API は、他の構成では期待通り動作しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="6017e-110">The feature APIs may not work as expected in other configurations.</span></span> <span data-ttu-id="6017e-111">構成方法を学習するプラットフォームを選択します。</span><span class="sxs-lookup"><span data-stu-id="6017e-111">Select the platform you'd like to learn how to configure.</span></span>

### <a name="windows"></a>[<span data-ttu-id="6017e-112">Windows</span><span class="sxs-lookup"><span data-stu-id="6017e-112">Windows</span></span>](#tab/windows)

#### <a name="shared-folders"></a><span data-ttu-id="6017e-113">共有フォルダー</span><span class="sxs-lookup"><span data-stu-id="6017e-113">Shared folders</span></span>

<span data-ttu-id="6017e-114">メールボックスの所有者は、まず代理人 [へのアクセス権を提供する必要があります](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="6017e-114">The mailbox owner must first [provide access to a delegate](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="6017e-115">代理人は、「他のユーザーのメールと予定表アイテムを管理する」の「自分のプロファイルに他のユーザーのメールボックスを追加する」セクションに記載されている手順に従う [必要があります](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)。</span><span class="sxs-lookup"><span data-stu-id="6017e-115">The delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="6017e-116">共有メールボックス (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="6017e-116">Shared mailboxes (preview)</span></span>

<span data-ttu-id="6017e-117">Exchange管理者は、アクセスするユーザーのセットの共有メールボックスを作成および管理できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-117">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="6017e-118">現時点では[、Exchange Online](/exchange/collaboration-exo/shared-mailboxes)この機能でサポートされている唯一のサーバー バージョンです。</span><span class="sxs-lookup"><span data-stu-id="6017e-118">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="6017e-119">"自動Exchange Server" と呼ばれる機能が既定でオンになっています。つまり、Outlook が閉[](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)じて再び開いた後、共有メールボックスはユーザーの Outlook アプリに自動的に表示されます。</span><span class="sxs-lookup"><span data-stu-id="6017e-119">An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened.</span></span> <span data-ttu-id="6017e-120">ただし、管理者が自動マップをオフにした場合、ユーザーは記事「開く」の「Outlook に共有メールボックスを追加して、Outlook で共有メールボックスを使用する」に示されている手動の手順[に従う必要があります](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd)。</span><span class="sxs-lookup"><span data-stu-id="6017e-120">However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span></span>

> [!WARNING]
> <span data-ttu-id="6017e-121">パスワード **を** 使用して共有メールボックスにサインインしない。</span><span class="sxs-lookup"><span data-stu-id="6017e-121">Do **NOT** sign into the shared mailbox with a password.</span></span> <span data-ttu-id="6017e-122">この場合、機能 API は機能しません。</span><span class="sxs-lookup"><span data-stu-id="6017e-122">The feature APIs won't work in that case.</span></span>

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="6017e-123">Web ブラウザー - モダン Outlook</span><span class="sxs-lookup"><span data-stu-id="6017e-123">Web browser - modern Outlook</span></span>](#tab/modern)

#### <a name="shared-folders"></a><span data-ttu-id="6017e-124">共有フォルダー</span><span class="sxs-lookup"><span data-stu-id="6017e-124">Shared folders</span></span>

<span data-ttu-id="6017e-125">メールボックスの所有者は、まずメールボックス フォルダー [のアクセス許可を](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) 更新して代理人へのアクセスを提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6017e-125">The mailbox owner must first [provide access to a delegate](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) by updating the mailbox folder permissions.</span></span> <span data-ttu-id="6017e-126">代理人は、記事「他のユーザーのメールボックスにアクセスする」の「Outlook Web App のフォルダー リストに他のユーザーのメールボックスを追加する」に記載されている手順に従う[必要があります](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081)。</span><span class="sxs-lookup"><span data-stu-id="6017e-126">The delegate must then follow the instructions outlined in the "Add another person’s mailbox to your folder list in Outlook Web App" section of the article [Access another person's mailbox](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="6017e-127">共有メールボックス (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="6017e-127">Shared mailboxes (preview)</span></span>

<span data-ttu-id="6017e-128">Exchange管理者は、アクセスするユーザーのセットの共有メールボックスを作成および管理できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-128">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="6017e-129">現時点では[、Exchange Online](/exchange/collaboration-exo/shared-mailboxes)この機能でサポートされている唯一のサーバー バージョンです。</span><span class="sxs-lookup"><span data-stu-id="6017e-129">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="6017e-130">アクセスを受け取った後、共有メールボックス ユーザーは、「Open」の「共有メールボックスを追加してプライマリ メールボックスの下に表示する」セクションに示されている手順に従って、Outlook on the web で共有メールボックスを使用する[必要があります](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207)。</span><span class="sxs-lookup"><span data-stu-id="6017e-130">After receiving access, a shared mailbox user must follow the steps outlined in the "Add the shared mailbox so it displays under your primary mailbox" section of the article [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span></span>

> [!WARNING]
> <span data-ttu-id="6017e-131">" **別の** メールボックスを開く" などの他のオプションを使用しない。</span><span class="sxs-lookup"><span data-stu-id="6017e-131">Do **NOT** use other options like "Open another mailbox".</span></span> <span data-ttu-id="6017e-132">機能 API が正しく動作しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="6017e-132">The feature APIs may not work properly then.</span></span>

---

<span data-ttu-id="6017e-133">アドインが一般的にアクティブ化する場所とアクティブ化しない場所の詳細については、「Outlook[](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)アドインの概要」ページの「アドインで使用可能なメールボックス アイテム」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="6017e-133">To learn more about where add-ins do and do not activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.</span></span>

## <a name="supported-permissions"></a><span data-ttu-id="6017e-134">サポートされているアクセス許可</span><span class="sxs-lookup"><span data-stu-id="6017e-134">Supported permissions</span></span>

<span data-ttu-id="6017e-135">次の表では、代理人および共有メールボックス ユーザー Office JavaScript API でサポートされるアクセス許可について説明します。</span><span class="sxs-lookup"><span data-stu-id="6017e-135">The following table describes the permissions that the Office JavaScript API supports for delegates and shared mailbox users.</span></span>

|<span data-ttu-id="6017e-136">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="6017e-136">Permission</span></span>|<span data-ttu-id="6017e-137">値</span><span class="sxs-lookup"><span data-stu-id="6017e-137">Value</span></span>|<span data-ttu-id="6017e-138">説明</span><span class="sxs-lookup"><span data-stu-id="6017e-138">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="6017e-139">読み取り</span><span class="sxs-lookup"><span data-stu-id="6017e-139">Read</span></span>|<span data-ttu-id="6017e-140">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="6017e-140">1 (000001)</span></span>|<span data-ttu-id="6017e-141">アイテムを読み取り可能。</span><span class="sxs-lookup"><span data-stu-id="6017e-141">Can read items.</span></span>|
|<span data-ttu-id="6017e-142">書き込み</span><span class="sxs-lookup"><span data-stu-id="6017e-142">Write</span></span>|<span data-ttu-id="6017e-143">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="6017e-143">2 (000010)</span></span>|<span data-ttu-id="6017e-144">アイテムを作成できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-144">Can create items.</span></span>|
|<span data-ttu-id="6017e-145">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="6017e-145">DeleteOwn</span></span>|<span data-ttu-id="6017e-146">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="6017e-146">4 (000100)</span></span>|<span data-ttu-id="6017e-147">作成したアイテムのみを削除できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-147">Can delete only the items they created.</span></span>|
|<span data-ttu-id="6017e-148">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="6017e-148">DeleteAll</span></span>|<span data-ttu-id="6017e-149">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="6017e-149">8 (001000)</span></span>|<span data-ttu-id="6017e-150">任意のアイテムを削除できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-150">Can delete any items.</span></span>|
|<span data-ttu-id="6017e-151">EditOwn</span><span class="sxs-lookup"><span data-stu-id="6017e-151">EditOwn</span></span>|<span data-ttu-id="6017e-152">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="6017e-152">16 (010000)</span></span>|<span data-ttu-id="6017e-153">作成したアイテムのみを編集できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-153">Can edit only the items they created.</span></span>|
|<span data-ttu-id="6017e-154">EditAll</span><span class="sxs-lookup"><span data-stu-id="6017e-154">EditAll</span></span>|<span data-ttu-id="6017e-155">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="6017e-155">32 (100000)</span></span>|<span data-ttu-id="6017e-156">任意のアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-156">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="6017e-157">現在、API は既存のアクセス許可の取得をサポートしていますが、アクセス許可を設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="6017e-157">Currently the API supports getting existing permissions, but not setting permissions.</span></span>

<span data-ttu-id="6017e-158">[DelegatePermissions オブジェクト](/javascript/api/outlook/office.mailboxenums.delegatepermissions)は、アクセス許可を示すためにビットマスクを使用して実装されます。</span><span class="sxs-lookup"><span data-stu-id="6017e-158">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the permissions.</span></span> <span data-ttu-id="6017e-159">ビットマスク内の各位置は、特定のアクセス許可を表し、それが設定されている場合、 `1` ユーザーはそれぞれのアクセス許可を持っています。</span><span class="sxs-lookup"><span data-stu-id="6017e-159">Each position in the bitmask represents a particular permission and if it's set to `1` then the user has the respective permission.</span></span> <span data-ttu-id="6017e-160">たとえば、右側の 2 番目のビットが次の場合 `1` 、ユーザーは書き込みアクセス **許可を持** つ。</span><span class="sxs-lookup"><span data-stu-id="6017e-160">For example, if the second bit from the right is `1`, then the user has **Write** permission.</span></span> <span data-ttu-id="6017e-161">特定のアクセス許可を確認する方法の例については、後の「[](#perform-an-operation-as-delegate-or-shared-mailbox-user)代理人または共有メールボックス ユーザーとして操作を実行する」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="6017e-161">You can see an example of how to check for a specific permission in the [Perform an operation as delegate or shared mailbox user](#perform-an-operation-as-delegate-or-shared-mailbox-user) section later in this article.</span></span>

## <a name="sync-across-shared-folder-clients"></a><span data-ttu-id="6017e-162">共有フォルダー クライアント間の同期</span><span class="sxs-lookup"><span data-stu-id="6017e-162">Sync across shared folder clients</span></span>

<span data-ttu-id="6017e-163">所有者のメールボックスに対する代理人の更新は、通常、すぐにメールボックス間で同期されます。</span><span class="sxs-lookup"><span data-stu-id="6017e-163">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="6017e-164">ただし、REST または Exchange Web サービス (EWS) 操作を使用してアイテムに拡張プロパティを設定した場合、このような変更は同期に数時間かかる可能性があります。このような遅延を回避するには[、CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトと関連する API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6017e-164">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="6017e-165">詳細については、「Get [and](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) set metadata in the Outlook」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6017e-165">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6017e-166">代理人シナリオでは、API によって現在提供されているトークンで EWS をoffice.jsできません。</span><span class="sxs-lookup"><span data-stu-id="6017e-166">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="6017e-167">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="6017e-167">Configure the manifest</span></span>

<span data-ttu-id="6017e-168">アドインで共有フォルダーと共有メールボックスのシナリオを有効にするには、親要素の下のマニフェストで [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) 要素を設定 `true` する必要があります `DesktopFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="6017e-168">To enable shared folders and shared mailbox scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="6017e-169">現時点では、他のフォーム ファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6017e-169">At present, other form factors are not supported.</span></span>

<span data-ttu-id="6017e-170">代理人からの REST 呼び出しをサポートするには、マニフェストの [Permissions](../reference/manifest/permissions.md) ノードをに設定します `ReadWriteMailbox` 。</span><span class="sxs-lookup"><span data-stu-id="6017e-170">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="6017e-171">次の例は、 `SupportsSharedFolders` マニフェストのセクションに `true` 設定された要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="6017e-171">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a><span data-ttu-id="6017e-172">代理人または共有メールボックス ユーザーとして操作を実行する</span><span class="sxs-lookup"><span data-stu-id="6017e-172">Perform an operation as delegate or shared mailbox user</span></span>

<span data-ttu-id="6017e-173">[item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出すことによって、作成モードまたは読み取りモードでアイテムの共有プロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-173">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="6017e-174">これにより、現在ユーザーのアクセス許可、所有者の電子メール アドレス、REST API の基本 URL、およびターゲット メールボックスを提供する [SharedProperties](/javascript/api/outlook/office.sharedproperties) オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="6017e-174">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the user's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="6017e-175">次の例は、メッセージまたは予定の共有プロパティを取得し、代理人または共有メールボックス ユーザーが **書** き込みアクセス許可を持つか確認し、REST 呼び出しを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="6017e-175">The following example shows how to get the shared properties of a message or appointment, check if the delegate or shared mailbox user has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="6017e-176">代理人として REST を使用して、アイテムまたはグループの投稿にOutlookメッセージのOutlook[を取得できます](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。</span><span class="sxs-lookup"><span data-stu-id="6017e-176">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="6017e-177">共有アイテムと非共有アイテムの REST の呼び出しを処理する</span><span class="sxs-lookup"><span data-stu-id="6017e-177">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="6017e-178">アイテムに対して REST 操作を呼び出す場合は、アイテムが共有されるかどうかに関して、API を使用してアイテムが共有 `getSharedPropertiesAsync` されているかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-178">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="6017e-179">その後、適切なオブジェクトを使用して操作の REST URL を作成できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-179">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

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

## <a name="limitations"></a><span data-ttu-id="6017e-180">制限事項</span><span class="sxs-lookup"><span data-stu-id="6017e-180">Limitations</span></span>

<span data-ttu-id="6017e-181">アドインのシナリオによっては、共有フォルダーや共有メールボックスの状況を処理する際に考慮すべきいくつかの制限があります。</span><span class="sxs-lookup"><span data-stu-id="6017e-181">Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="6017e-182">メッセージ作成モード</span><span class="sxs-lookup"><span data-stu-id="6017e-182">Message Compose mode</span></span>

<span data-ttu-id="6017e-183">メッセージ作成モードでは、次の条件を満たしていない限り、Outlook on the web または Windows で[getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_)はサポートされません。</span><span class="sxs-lookup"><span data-stu-id="6017e-183">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or on Windows unless the following conditions are met.</span></span>

<span data-ttu-id="6017e-184">a.</span><span class="sxs-lookup"><span data-stu-id="6017e-184">a.</span></span> <span data-ttu-id="6017e-185">**アクセス/共有フォルダーの委任**</span><span class="sxs-lookup"><span data-stu-id="6017e-185">**Delegate access/Shared folders**</span></span>

1. <span data-ttu-id="6017e-186">メールボックスの所有者がメッセージを開始します。</span><span class="sxs-lookup"><span data-stu-id="6017e-186">The mailbox owner starts a message.</span></span> <span data-ttu-id="6017e-187">これは、新しいメッセージ、返信、または転送を指定できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-187">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="6017e-188">メッセージを保存し、そのメッセージを自分の **下** 書きフォルダーから代理人と共有するフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="6017e-188">They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.</span></span>
1. <span data-ttu-id="6017e-189">代理人は、共有フォルダーから下書きを開き、作成を続行します。</span><span class="sxs-lookup"><span data-stu-id="6017e-189">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="6017e-190">b.</span><span class="sxs-lookup"><span data-stu-id="6017e-190">b.</span></span> <span data-ttu-id="6017e-191">**共有メールボックス**</span><span class="sxs-lookup"><span data-stu-id="6017e-191">**Shared mailbox**</span></span>

1. <span data-ttu-id="6017e-192">共有メールボックス ユーザーがメッセージを開始します。</span><span class="sxs-lookup"><span data-stu-id="6017e-192">A shared mailbox user starts a message.</span></span> <span data-ttu-id="6017e-193">これは、新しいメッセージ、返信、または転送を指定できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-193">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="6017e-194">メッセージを保存し、そのメッセージを自分の **下** 書きフォルダーから共有メールボックス内のフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="6017e-194">They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.</span></span>
1. <span data-ttu-id="6017e-195">別の共有メールボックス ユーザーが共有メールボックスから下書きを開き、作成を続行します。</span><span class="sxs-lookup"><span data-stu-id="6017e-195">Another shared mailbox user opens the draft from the shared mailbox then continues composing.</span></span>

<span data-ttu-id="6017e-196">このメッセージは共有コンテキストに追加され、これらの共有シナリオをサポートするアドインはアイテムの共有プロパティを取得できます。</span><span class="sxs-lookup"><span data-stu-id="6017e-196">The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.</span></span> <span data-ttu-id="6017e-197">メッセージが送信された後、通常は送信者の [送信アイテム] **フォルダーにあります** 。</span><span class="sxs-lookup"><span data-stu-id="6017e-197">After the message has been sent, it's usually found in the sender's **Sent Items** folder.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="6017e-198">REST と EWS</span><span class="sxs-lookup"><span data-stu-id="6017e-198">REST and EWS</span></span>

<span data-ttu-id="6017e-199">アドインは REST を使用できます。また、アドインのアクセス許可を設定して、所有者のメールボックスまたは共有メールボックスへの REST アクセスを有効にする必要 `ReadWriteMailbox` があります。</span><span class="sxs-lookup"><span data-stu-id="6017e-199">Your add-in can use REST and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox or to the shared mailbox as applicable.</span></span> <span data-ttu-id="6017e-200">EWS はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6017e-200">EWS is not supported.</span></span>

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a><span data-ttu-id="6017e-201">アドレス一覧から非表示のユーザーまたは共有メールボックス</span><span class="sxs-lookup"><span data-stu-id="6017e-201">User or shared mailbox hidden from an address list</span></span>

<span data-ttu-id="6017e-202">管理者がグローバル アドレス一覧 (GAL) などのアドレス一覧からユーザーまたは共有メールボックス のアドレスを非表示にした場合、メールボックス レポートで開いた影響を受けるメール アイテムは `Office.context.mailbox.item` null として開きます。</span><span class="sxs-lookup"><span data-stu-id="6017e-202">If an admin hid a user or shared mailbox address from an address list like the global address list (GAL), affected mail items opened in the mailbox report `Office.context.mailbox.item` as null.</span></span> <span data-ttu-id="6017e-203">たとえば、ユーザーが GAL から非表示の共有メールボックスでメール アイテムを開くと、そのメール アイテムは `Office.context.mailbox.item` null になります。</span><span class="sxs-lookup"><span data-stu-id="6017e-203">For example, if the user opens a mail item in a shared mailbox that's hidden from the GAL, `Office.context.mailbox.item` representing that mail item is null.</span></span>

## <a name="see-also"></a><span data-ttu-id="6017e-204">関連項目</span><span class="sxs-lookup"><span data-stu-id="6017e-204">See also</span></span>

- [<span data-ttu-id="6017e-205">自分のメールと予定表の管理を他のユーザーに許可する</span><span class="sxs-lookup"><span data-stu-id="6017e-205">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="6017e-206">カレンダーの共有Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="6017e-206">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="6017e-207">共有メールボックスをユーザーに追加Outlook</span><span class="sxs-lookup"><span data-stu-id="6017e-207">Add a shared mailbox to Outlook</span></span>](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [<span data-ttu-id="6017e-208">マニフェスト要素を順序付けする方法</span><span class="sxs-lookup"><span data-stu-id="6017e-208">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="6017e-209">[マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="6017e-209">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="6017e-210">JavaScript ビット演算子</span><span class="sxs-lookup"><span data-stu-id="6017e-210">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)