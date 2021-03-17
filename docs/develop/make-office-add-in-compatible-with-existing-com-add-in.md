---
title: Office アドインに既存の COM アドインとの互換性をもたせる
description: アドインと同等の COM アドインOffice互換性を有効にする。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836855"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="e1dd8-103">Office アドインに既存の COM アドインとの互換性をもたせる</span><span class="sxs-lookup"><span data-stu-id="e1dd8-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="e1dd8-104">既存の COM アドインがある場合は、Office アドインで同等の機能を構築して、web や Mac 上の Office などの他のプラットフォームでソリューションを実行できます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="e1dd8-105">場合によっては、Office COM アドインで使用可能なすべての機能を提供できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="e1dd8-106">このような状況では、COM アドインは、アドインが提供できる対応する機能よりも、Windows でのユーザー エクスペリエンスOffice向上することがあります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="e1dd8-107">Office アドインを構成して、同等の COM アドインが既にユーザーのコンピューターにインストールされている場合、Windows 上の Office が Office アドインではなく COM アドインを実行します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="e1dd8-108">COM アドインは、Office がユーザーのコンピューターにインストールされている COM アドインと Office アドインの間でシームレスに切り替わるため、「同等」と呼ばれる。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="e1dd8-109">この機能は、Microsoft 365 サブスクリプションに接続されている場合、次のプラットフォームでサポートされます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="e1dd8-110">Web 上の Excel、Word、および PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e1dd8-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="e1dd8-111">Windows 上の Excel、Word、および PowerPoint (バージョン 1904 以降)</span><span class="sxs-lookup"><span data-stu-id="e1dd8-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="e1dd8-112">Mac 上の Excel、Word、および PowerPoint (バージョン 13.329 以降)</span><span class="sxs-lookup"><span data-stu-id="e1dd8-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>
> - <span data-ttu-id="e1dd8-113">Outlook on Windows (バージョン 2102 以降)</span><span class="sxs-lookup"><span data-stu-id="e1dd8-113">Outlook on Windows (version 2102 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="e1dd8-114">同等の COM アドインを指定する</span><span class="sxs-lookup"><span data-stu-id="e1dd8-114">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="e1dd8-115">マニフェスト</span><span class="sxs-lookup"><span data-stu-id="e1dd8-115">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e1dd8-116">Excel、PowerPoint、および Word に適用されます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-116">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="e1dd8-117">Outlook のサポートは近日公開予定です。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-117">Outlook support coming soon.</span></span>

<span data-ttu-id="e1dd8-118">Office アドインと COM アドイン間の互換性を有効にするには、Office アドインのマニフェストで同等の COM アドインを[](add-in-manifests.md)識別します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-118">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="e1dd8-119">次Office Windows では、両方がインストールされている場合は、Officeアドインの代わりに COM アドインが使用されます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-119">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="e1dd8-120">次の例は、COM アドインを同等のアドインとして指定するマニフェストの部分を示しています。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-120">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="e1dd8-121">要素の値は COM アドインを識別し `ProgId` [、EquivalentAddins](../reference/manifest/equivalentaddins.md) 要素は終了タグの直前に配置する必要 `VersionOverrides` があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-121">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="e1dd8-122">COM アドインと XLL UDF の互換性については、「カスタム関数を XLL ユーザー定義関数と互換性のあるものにする [」を参照してください](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-122">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="e1dd8-123">グループ ポリシー</span><span class="sxs-lookup"><span data-stu-id="e1dd8-123">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e1dd8-124">Outlook にのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-124">Applies to Outlook only.</span></span>

<span data-ttu-id="e1dd8-125">Outlook Web アドインと COM/VSTO アドインの互換性を宣言するには、グループ ポリシーで同等の COM アドインを識別し、ユーザーのコンピューターで構成することにより、同等の COM または **VSTO** アドインがインストールされている Outlook Web アドインを非アクティブ化します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-125">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="e1dd8-126">次に、Outlook on Windows では、両方がインストールされている場合は、Web アドインの代わりに COM アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-126">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="e1dd8-127">ツールの [インストール手順に](https://www.microsoft.com/download/details.aspx?id=49030)注意を払って、最新の管理用テンプレート ツール **をダウンロードします**。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-127">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="e1dd8-128">ローカル グループ ポリシー エディター **(gpedit.msc) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-128">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="e1dd8-129">[ユーザー **構成]**  >  **[管理用テンプレート**]   >  **[Microsoft Outlook 2016**  >  **その他] に移動します**。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-129">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="e1dd8-130">同等の **COM または VSTO** アドインがインストールされている Outlook Web アドインを非アクティブ化する設定を選択します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-130">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="e1dd8-131">リンクを開き、ポリシー設定を編集します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-131">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="e1dd8-132">ダイアログの **Outlook Web アドインで非アクティブ化するには、次の操作を行います**。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-132">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="e1dd8-133">[ **値の名前]** を `Id` Web アドインのマニフェストで見つかった名前に設定します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-133">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="e1dd8-134">**重要**: *中かっこ* をエントリの周囲 `{}` に追加しない。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-134">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="e1dd8-135">Value **を** 同等 `ProgId` の COM/VSTO アドインの値に設定します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-135">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="e1dd8-136">**[OK] を** 選択して更新プログラムを有効にします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-136">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="e1dd8-137">![ダイアログ "非アクティブ化する Outlook Web アドイン" を示すスクリーンショット](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="e1dd8-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate"](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="e1dd8-138">ユーザーと同等の動作</span><span class="sxs-lookup"><span data-stu-id="e1dd8-138">Equivalent behavior for users</span></span>

<span data-ttu-id="e1dd8-139">同等の [COM](#specify-an-equivalent-com-add-in)アドインを指定すると、windows 上の Office は、同等の COM アドインがインストールされている場合、Office アドインのユーザー インターフェイス (UI) は表示されません。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-139">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="e1dd8-140">Officeアドインのリボン ボタンのみを非表示Officeし、インストールを妨げる必要があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-140">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="e1dd8-141">したがって、Officeアドインは UI 内の次の場所に表示されます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-141">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="e1dd8-142">[ **自分のアドイン] の下**</span><span class="sxs-lookup"><span data-stu-id="e1dd8-142">Under **My add-ins**</span></span>
- <span data-ttu-id="e1dd8-143">リボン マネージャーのエントリとして (Excel、Word、および PowerPoint のみ)</span><span class="sxs-lookup"><span data-stu-id="e1dd8-143">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="e1dd8-144">マニフェストで同等の COM アドインを指定すると、web 上や Mac 上の Officeなどの他のプラットフォームには影響しません。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-144">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="e1dd8-145">次のシナリオでは、ユーザーがアドインを取得する方法に応Office説明します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-145">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="e1dd8-146">AppSource によるアドインOffice取得</span><span class="sxs-lookup"><span data-stu-id="e1dd8-146">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="e1dd8-147">ユーザーが AppSource から Officeアドインを取得し、同等の COM アドインが既にインストールされている場合は、次Officeします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-147">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="e1dd8-148">アドインOfficeインストールします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-148">Install the Office Add-in.</span></span>
2. <span data-ttu-id="e1dd8-149">リボンでOfficeアドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-149">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="e1dd8-150">COM アドイン リボン ボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-150">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="e1dd8-151">アドインのOffice展開</span><span class="sxs-lookup"><span data-stu-id="e1dd8-151">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="e1dd8-152">管理者が集中展開を使用して Office アドインをテナントに展開し、同等の COM アドインが既にインストールされている場合、ユーザーは変更を表示する前に Office を再起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-152">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="e1dd8-153">再起動Office、次のコマンドが実行されます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-153">After Office restarts, it will:</span></span>

1. <span data-ttu-id="e1dd8-154">アドインOfficeインストールします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-154">Install the Office Add-in.</span></span>
2. <span data-ttu-id="e1dd8-155">リボンでOfficeアドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-155">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="e1dd8-156">COM アドイン リボン ボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-156">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="e1dd8-157">埋め込みアドインと共有Officeドキュメント</span><span class="sxs-lookup"><span data-stu-id="e1dd8-157">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="e1dd8-158">ユーザーが COM アドインをインストールし、埋め込み Office アドインを含む共有ドキュメントを取得した場合、そのユーザーがドキュメントを開いた場合、次のOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-158">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="e1dd8-159">ユーザーにアドインを信頼Office求めるメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-159">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="e1dd8-160">信頼できる場合は、Officeアドインがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-160">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="e1dd8-161">リボンでOfficeアドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-161">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="e1dd8-162">その他の COM アドインの動作</span><span class="sxs-lookup"><span data-stu-id="e1dd8-162">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="e1dd8-163">Excel、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="e1dd8-163">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="e1dd8-164">ユーザーが同等の COM アドインをアンインストールした場合は、Windows Officeアドイン UI Office復元します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-164">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="e1dd8-165">カスタム アドインに同等の COM アドインを指定したOffice、Officeの更新プログラムの処理Office停止します。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-165">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="e1dd8-166">アドインの最新の更新プログラムOffice、ユーザーはまず COM アドインをアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-166">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="e1dd8-167">Outlook</span><span class="sxs-lookup"><span data-stu-id="e1dd8-167">Outlook</span></span>

<span data-ttu-id="e1dd8-168">対応する Web アドインを無効にするには、Outlook の起動時に COM/VSTO アドインを接続する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-168">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="e1dd8-169">その後の Outlook セッション中に COM/VSTO アドインが切断された場合、Outlook が再起動されるまで、Web アドインは無効なままである可能性があります。</span><span class="sxs-lookup"><span data-stu-id="e1dd8-169">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="e1dd8-170">関連項目</span><span class="sxs-lookup"><span data-stu-id="e1dd8-170">See also</span></span>

- [<span data-ttu-id="e1dd8-171">カスタム関数を XLL ユーザー定義関数と互換性のあるものにする</span><span class="sxs-lookup"><span data-stu-id="e1dd8-171">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
