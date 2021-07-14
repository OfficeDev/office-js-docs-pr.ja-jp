---
title: Office アドインに既存の COM アドインとの互換性をもたせる
description: アドインと同等の COM アドインOffice互換性を有効にする。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 85e5d8cc06aa599862c92b59a26c744f28ca2d22
ms.sourcegitcommit: 95fc1fc8a0dbe8fc94f0ea647836b51cc7f8601d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/14/2021
ms.locfileid: "53418686"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="c78ca-103">Office アドインに既存の COM アドインとの互換性をもたせる</span><span class="sxs-lookup"><span data-stu-id="c78ca-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="c78ca-104">既存の COM アドインがある場合は、Office アドインで同等の機能を構築して、Office on the web や Mac などの他のプラットフォームでソリューションを実行できます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="c78ca-105">場合によっては、Officeアドインが、対応する COM アドインで使用できるすべての機能を提供できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="c78ca-106">このような状況では、COM アドインは、アドインが提供できる対応するWindowsよりも、Officeユーザー エクスペリエンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="c78ca-107">Office アドインを構成して、同等の COM アドインが既にユーザーのコンピューターにインストールされている場合、Windows の Office が Office アドインの代わりに COM アドインを実行します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="c78ca-108">COM アドインは、Office がユーザーのコンピューターにインストールされているに従って、COM アドインと Office アドインの間でシームレスに切り替わるため、「同等」と呼ばれる。</span><span class="sxs-lookup"><span data-stu-id="c78ca-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="c78ca-109">この機能は、サブスクリプションに接続されている場合、次のプラットフォームとアプリケーションMicrosoft 365されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-109">This feature is supported by the following platform and applications, when connected to a Microsoft 365 subscription.</span></span> <span data-ttu-id="c78ca-110">COM アドインは他のプラットフォームにインストールできないので、これらのプラットフォームでは、この記事で後で説明する manifest 要素 `EquivalentAddins` は無視されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-110">COM add-ins cannot be installed on any other platform, so on those platforms the manifest element that is discussed later in this article, `EquivalentAddins`, is ignored.</span></span>
>
> - <span data-ttu-id="c78ca-111">Excel、Word、および PowerPoint (Windows 1904 以降)</span><span class="sxs-lookup"><span data-stu-id="c78ca-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="c78ca-112">同等の COM アドインを指定する</span><span class="sxs-lookup"><span data-stu-id="c78ca-112">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="c78ca-113">マニフェスト</span><span class="sxs-lookup"><span data-stu-id="c78ca-113">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c78ca-114">Word、Excel、PowerPointに適用されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-114">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="c78ca-115">Outlookサポートが近日公開されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-115">Outlook support coming soon.</span></span>

<span data-ttu-id="c78ca-116">Office アドインと COM アドイン間の互換性を有効にするには、Office アドインのマニフェストで同等の COM アドインを[](add-in-manifests.md)識別します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-116">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="c78ca-117">次Office、Windows両方がインストールされている場合は、Officeアドインではなく COM アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-117">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="c78ca-118">次の例は、COM アドインを同等のアドインとして指定するマニフェストの部分を示しています。</span><span class="sxs-lookup"><span data-stu-id="c78ca-118">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="c78ca-119">要素の値は COM アドインを識別し `ProgId` [、EquivalentAddins](../reference/manifest/equivalentaddins.md) 要素は終了タグの直前に配置する必要 `VersionOverrides` があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-119">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

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
> <span data-ttu-id="c78ca-120">COM アドインと XLL UDF の互換性については、「カスタム関数を XLL ユーザー定義関数と互換性のあるものにする [」を参照してください](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="c78ca-120">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="c78ca-121">グループ ポリシー</span><span class="sxs-lookup"><span data-stu-id="c78ca-121">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c78ca-122">ユーザーにのみOutlook適用されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-122">Applies to Outlook only.</span></span>

<span data-ttu-id="c78ca-123">Outlook Web アドインと COM/VSTO アドイン間の互換性を宣言するには、グループ ポリシー [非アクティブ化] Outlook Web アドインの同等の COM アドインまたは **VSTO** アドインをユーザーのコンピューターで構成してインストールする同等の COM アドインを識別します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-123">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="c78ca-124">次Outlook、Windowsがインストールされている場合、Web アドインの代わりに COM アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-124">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="c78ca-125">ツールの [インストール手順に](https://www.microsoft.com/download/details.aspx?id=49030)注意を払って、最新の管理用テンプレート ツール **をダウンロードします**。</span><span class="sxs-lookup"><span data-stu-id="c78ca-125">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="c78ca-126">ローカル グループ ポリシー エディター **(gpedit.msc) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="c78ca-126">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="c78ca-127">[ユーザー **構成] [**  >  **管理用テンプレート**   >  **] [Microsoft Outlook 2016**  >  **その他] に移動します**。</span><span class="sxs-lookup"><span data-stu-id="c78ca-127">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="c78ca-128">同等の COM または Outlookがインストールされている Web アドインを非アクティブ化するVSTO **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="c78ca-128">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="c78ca-129">リンクを開き、ポリシー設定を編集します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-129">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="c78ca-130">ダイアログ ボックスで **、Outlookを非アクティブ化します**。</span><span class="sxs-lookup"><span data-stu-id="c78ca-130">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="c78ca-131">[ **値の名前]** を `Id` Web アドインのマニフェストで見つかった名前に設定します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-131">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="c78ca-132">**重要**: *中かっこ* をエントリの周囲 `{}` に追加しない。</span><span class="sxs-lookup"><span data-stu-id="c78ca-132">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="c78ca-133">Value **を** 同等 `ProgId` の COM/VSTOに設定します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-133">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="c78ca-134">**[OK] を** 選択して更新プログラムを有効にします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-134">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="c78ca-135">!["非アクティブ化する web Outlookを表示する" ダイアログを示すスクリーンショット。](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="c78ca-135">![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="c78ca-136">ユーザーと同等の動作</span><span class="sxs-lookup"><span data-stu-id="c78ca-136">Equivalent behavior for users</span></span>

<span data-ttu-id="c78ca-137">同等の[COM](#specify-an-equivalent-com-add-in)アドインを指定すると、Windows の Office は、同等の COM アドインがインストールされている場合、Office アドインのユーザー インターフェイス (UI) は表示されません。</span><span class="sxs-lookup"><span data-stu-id="c78ca-137">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="c78ca-138">Officeアドインのリボン ボタンのみを非表示にし、インストールOffice防ぐ必要があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-138">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="c78ca-139">したがって、Officeアドインは UI 内の次の場所に表示されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-139">Therefore your Office Add-in will still appear in the following locations within the UI.</span></span>

- <span data-ttu-id="c78ca-140">[ **自分のアドイン] の下**</span><span class="sxs-lookup"><span data-stu-id="c78ca-140">Under **My add-ins**</span></span>
- <span data-ttu-id="c78ca-141">リボン マネージャーのエントリとして (Excel、Word、およびPowerPointのみ)</span><span class="sxs-lookup"><span data-stu-id="c78ca-141">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="c78ca-142">マニフェストで同等の COM アドインを指定すると、他のプラットフォーム (Office on the web Mac など) には影響しません。</span><span class="sxs-lookup"><span data-stu-id="c78ca-142">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="c78ca-143">次のシナリオでは、ユーザーがアドインを取得する方法に応Office説明します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-143">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="c78ca-144">AppSource によるアドインOffice取得</span><span class="sxs-lookup"><span data-stu-id="c78ca-144">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="c78ca-145">ユーザーが AppSource から Officeアドインを取得し、同等の COM アドインが既にインストールされている場合は、次Officeします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-145">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="c78ca-146">アドインOfficeインストールします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-146">Install the Office Add-in.</span></span>
2. <span data-ttu-id="c78ca-147">リボンでOfficeアドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-147">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="c78ca-148">COM アドイン リボン ボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-148">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="c78ca-149">アドインのOffice展開</span><span class="sxs-lookup"><span data-stu-id="c78ca-149">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="c78ca-150">管理者が集中展開を使用して Office アドインをテナントに展開し、同等の COM アドインが既にインストールされている場合、ユーザーは変更を表示する前に Office を再起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-150">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="c78ca-151">再起動Office、次のコマンドが実行されます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-151">After Office restarts, it will:</span></span>

1. <span data-ttu-id="c78ca-152">アドインOfficeインストールします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-152">Install the Office Add-in.</span></span>
2. <span data-ttu-id="c78ca-153">リボンでOfficeアドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-153">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="c78ca-154">COM アドイン リボン ボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-154">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="c78ca-155">埋め込みアドインと共有Officeドキュメント</span><span class="sxs-lookup"><span data-stu-id="c78ca-155">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="c78ca-156">ユーザーが COM アドインをインストールし、埋め込み Office アドインを含む共有ドキュメントを取得した場合、そのユーザーがドキュメントを開いた場合、次のOffice。</span><span class="sxs-lookup"><span data-stu-id="c78ca-156">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="c78ca-157">ユーザーにアドインを信頼Office求めるメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-157">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="c78ca-158">信頼できる場合は、Officeアドインがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="c78ca-158">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="c78ca-159">リボンでOfficeアドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="c78ca-159">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="c78ca-160">その他の COM アドインの動作</span><span class="sxs-lookup"><span data-stu-id="c78ca-160">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="c78ca-161">Excel、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="c78ca-161">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="c78ca-162">ユーザーが同等の COM アドインをアンインストールした場合は、OfficeをWindows、Office UI を復元します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-162">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="c78ca-163">カスタム アドインに同等の COM アドインを指定したOffice、Officeアドインの更新プログラムの処理Office停止します。</span><span class="sxs-lookup"><span data-stu-id="c78ca-163">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="c78ca-164">アドインの最新の更新プログラムOffice、ユーザーはまず COM アドインをアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-164">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="c78ca-165">Outlook</span><span class="sxs-lookup"><span data-stu-id="c78ca-165">Outlook</span></span>

<span data-ttu-id="c78ca-166">対応する web VSTOを無効にするために、Outlookを開始するときに、COM/Outlookアドインを接続する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-166">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="c78ca-167">その後の Outlook セッション中に COM/VSTO アドインが切断された場合、Web アドインは再起動するまで無効Outlook可能性があります。</span><span class="sxs-lookup"><span data-stu-id="c78ca-167">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="c78ca-168">関連項目</span><span class="sxs-lookup"><span data-stu-id="c78ca-168">See also</span></span>

- [<span data-ttu-id="c78ca-169">カスタム関数を XLL ユーザー定義関数と互換性のあるものにする</span><span class="sxs-lookup"><span data-stu-id="c78ca-169">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
