---
title: Office アドインの XML マニフェスト
description: ''
ms.date: 08/14/2019
localization_priority: Priority
ms.openlocfilehash: da8e865a78b666d4790df854403604cc03d6a47a
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477916"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="776d6-102">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="776d6-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="776d6-103">Office アドインの XML マニフェスト ファイルでは、エンド ユーザーが Office ドキュメントや Office アプリケーションにアドインをインストールして使用するときにアドインをアクティブ化する方法が記述されています。</span><span class="sxs-lookup"><span data-stu-id="776d6-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="776d6-104">このスキーマに基づいた XML マニフェスト ファイルを使用すると、Office アドインで次のことができます。</span><span class="sxs-lookup"><span data-stu-id="776d6-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="776d6-105">ID、バージョン、説明、表示名、および既定のロケールを指定することで、アプリ自体について説明する。</span><span class="sxs-lookup"><span data-stu-id="776d6-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="776d6-106">アドインのブランド化に使用するイメージと、Office リボンで[アドイン コマンド][]に使用する画像を指定する。</span><span class="sxs-lookup"><span data-stu-id="776d6-106">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="776d6-107">アドインを Office に統合する方法を指定する。アドインによって作成されるカスタム UI (リボンのボタンなど) の統合も含む。</span><span class="sxs-lookup"><span data-stu-id="776d6-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="776d6-108">コンテンツ アドインに必要な既定のサイズ、および Outlook アドインに必要な高さを指定する。</span><span class="sxs-lookup"><span data-stu-id="776d6-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="776d6-109">ドキュメントの読み取り、書き込みなど、Office アドインに必要なアクセス許可を宣言する。</span><span class="sxs-lookup"><span data-stu-id="776d6-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="776d6-110">Outlook アドインでは、アプリがアクティブ化されてメッセージ、予定、または会議出席依頼アイテムを操作するコンテキストを指定するルールを定義する。</span><span class="sxs-lookup"><span data-stu-id="776d6-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="776d6-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="776d6-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="776d6-113">必要な要素</span><span class="sxs-lookup"><span data-stu-id="776d6-113">Required elements</span></span>

<span data-ttu-id="776d6-114">次の表に、3 種類の Office アドインに必要な要素を示します。</span><span class="sxs-lookup"><span data-stu-id="776d6-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="776d6-115">親要素内で要素を表示する順序も決まっています。</span><span class="sxs-lookup"><span data-stu-id="776d6-115">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="776d6-116">詳細については、[マニフェスト要素の正しい順序を確認する方法](manifest-element-ordering.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="776d6-116">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="776d6-117">Office アドインの種類ごとの必要な要素</span><span class="sxs-lookup"><span data-stu-id="776d6-117">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="776d6-118">要素</span><span class="sxs-lookup"><span data-stu-id="776d6-118">Element</span></span>                                                                                      | <span data-ttu-id="776d6-119">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="776d6-119">Content</span></span> | <span data-ttu-id="776d6-120">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="776d6-120">Task pane</span></span> | <span data-ttu-id="776d6-121">Outlook</span><span class="sxs-lookup"><span data-stu-id="776d6-121">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="776d6-122">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="776d6-122">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="776d6-123">X</span><span class="sxs-lookup"><span data-stu-id="776d6-123">X</span></span>    |     <span data-ttu-id="776d6-124">X</span><span class="sxs-lookup"><span data-stu-id="776d6-124">X</span></span>     |    <span data-ttu-id="776d6-125">X</span><span class="sxs-lookup"><span data-stu-id="776d6-125">X</span></span>    |
| <span data-ttu-id="776d6-126">
  [Id][]</span><span class="sxs-lookup"><span data-stu-id="776d6-126">[Id][]</span></span>                                                                                       |    <span data-ttu-id="776d6-127">X</span><span class="sxs-lookup"><span data-stu-id="776d6-127">X</span></span>    |     <span data-ttu-id="776d6-128">X</span><span class="sxs-lookup"><span data-stu-id="776d6-128">X</span></span>     |    <span data-ttu-id="776d6-129">X</span><span class="sxs-lookup"><span data-stu-id="776d6-129">X</span></span>    |
| <span data-ttu-id="776d6-130">
  [Version][]</span><span class="sxs-lookup"><span data-stu-id="776d6-130">[Version][]</span></span>                                                                                  |    <span data-ttu-id="776d6-131">X</span><span class="sxs-lookup"><span data-stu-id="776d6-131">X</span></span>    |     <span data-ttu-id="776d6-132">X</span><span class="sxs-lookup"><span data-stu-id="776d6-132">X</span></span>     |    <span data-ttu-id="776d6-133">X</span><span class="sxs-lookup"><span data-stu-id="776d6-133">X</span></span>    |
| <span data-ttu-id="776d6-134">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="776d6-134">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="776d6-135">X</span><span class="sxs-lookup"><span data-stu-id="776d6-135">X</span></span>    |     <span data-ttu-id="776d6-136">X</span><span class="sxs-lookup"><span data-stu-id="776d6-136">X</span></span>     |    <span data-ttu-id="776d6-137">X</span><span class="sxs-lookup"><span data-stu-id="776d6-137">X</span></span>    |
| <span data-ttu-id="776d6-138">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="776d6-138">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="776d6-139">X</span><span class="sxs-lookup"><span data-stu-id="776d6-139">X</span></span>    |     <span data-ttu-id="776d6-140">X</span><span class="sxs-lookup"><span data-stu-id="776d6-140">X</span></span>     |    <span data-ttu-id="776d6-141">X</span><span class="sxs-lookup"><span data-stu-id="776d6-141">X</span></span>    |
| <span data-ttu-id="776d6-142">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="776d6-142">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="776d6-143">X</span><span class="sxs-lookup"><span data-stu-id="776d6-143">X</span></span>    |     <span data-ttu-id="776d6-144">X</span><span class="sxs-lookup"><span data-stu-id="776d6-144">X</span></span>     |    <span data-ttu-id="776d6-145">X</span><span class="sxs-lookup"><span data-stu-id="776d6-145">X</span></span>    |
| <span data-ttu-id="776d6-146">[Description][]</span><span class="sxs-lookup"><span data-stu-id="776d6-146">[Description][]</span></span>                                                                              |    <span data-ttu-id="776d6-147">X</span><span class="sxs-lookup"><span data-stu-id="776d6-147">X</span></span>    |     <span data-ttu-id="776d6-148">X</span><span class="sxs-lookup"><span data-stu-id="776d6-148">X</span></span>     |    <span data-ttu-id="776d6-149">X</span><span class="sxs-lookup"><span data-stu-id="776d6-149">X</span></span>    |
| <span data-ttu-id="776d6-150">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="776d6-150">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="776d6-151">X</span><span class="sxs-lookup"><span data-stu-id="776d6-151">X</span></span>    |     <span data-ttu-id="776d6-152">X</span><span class="sxs-lookup"><span data-stu-id="776d6-152">X</span></span>     |    <span data-ttu-id="776d6-153">X</span><span class="sxs-lookup"><span data-stu-id="776d6-153">X</span></span>    |
| <span data-ttu-id="776d6-154">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-154">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="776d6-155">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-155">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="776d6-156">X</span><span class="sxs-lookup"><span data-stu-id="776d6-156">X</span></span>    |     <span data-ttu-id="776d6-157">X</span><span class="sxs-lookup"><span data-stu-id="776d6-157">X</span></span>     |         |
| <span data-ttu-id="776d6-158">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-158">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="776d6-159">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-159">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="776d6-160">X</span><span class="sxs-lookup"><span data-stu-id="776d6-160">X</span></span>    |     <span data-ttu-id="776d6-161">X</span><span class="sxs-lookup"><span data-stu-id="776d6-161">X</span></span>     |         |
| <span data-ttu-id="776d6-162">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="776d6-162">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="776d6-163">X</span><span class="sxs-lookup"><span data-stu-id="776d6-163">X</span></span>    |
| <span data-ttu-id="776d6-164">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-164">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="776d6-165">X</span><span class="sxs-lookup"><span data-stu-id="776d6-165">X</span></span>    |
| <span data-ttu-id="776d6-166">
  [Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-166">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="776d6-167">
  [Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-167">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="776d6-168">
  [Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-168">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="776d6-169">X</span><span class="sxs-lookup"><span data-stu-id="776d6-169">X</span></span>    |     <span data-ttu-id="776d6-170">X</span><span class="sxs-lookup"><span data-stu-id="776d6-170">X</span></span>     |    <span data-ttu-id="776d6-171">X</span><span class="sxs-lookup"><span data-stu-id="776d6-171">X</span></span>    |
| <span data-ttu-id="776d6-172">
  [Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-172">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="776d6-173">
  [Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="776d6-173">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="776d6-174">X</span><span class="sxs-lookup"><span data-stu-id="776d6-174">X</span></span>    |
| <span data-ttu-id="776d6-175">[Requirements (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-175">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="776d6-176">X</span><span class="sxs-lookup"><span data-stu-id="776d6-176">X</span></span>    |
| <span data-ttu-id="776d6-177">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-177">[Set\*][]</span></span><br/><span data-ttu-id="776d6-178">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-178">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="776d6-179">X</span><span class="sxs-lookup"><span data-stu-id="776d6-179">X</span></span>    |
| <span data-ttu-id="776d6-180">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-180">[Form\*][]</span></span><br/><span data-ttu-id="776d6-181">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-181">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="776d6-182">X</span><span class="sxs-lookup"><span data-stu-id="776d6-182">X</span></span>    |
| <span data-ttu-id="776d6-183">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-183">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="776d6-184">X</span><span class="sxs-lookup"><span data-stu-id="776d6-184">X</span></span>    |     <span data-ttu-id="776d6-185">X</span><span class="sxs-lookup"><span data-stu-id="776d6-185">X</span></span>     |         |
| <span data-ttu-id="776d6-186">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="776d6-186">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="776d6-187">X</span><span class="sxs-lookup"><span data-stu-id="776d6-187">X</span></span>    |     <span data-ttu-id="776d6-188">X</span><span class="sxs-lookup"><span data-stu-id="776d6-188">X</span></span>     |         |

<span data-ttu-id="776d6-189">_\*Office アドイン マニフェスト スキーマ バージョン 1.1 で追加されました。_</span><span class="sxs-lookup"><span data-stu-id="776d6-189">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: /office/dev/add-ins/reference/manifest/officeapp
[id]: /office/dev/add-ins/reference/manifest/id
[version]: /office/dev/add-ins/reference/manifest/version
[providername]: /office/dev/add-ins/reference/manifest/providername
[defaultlocale]: /office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: /office/dev/add-ins/reference/manifest/displayname
[description]: /office/dev/add-ins/reference/manifest/description
[iconurl]: /office/dev/add-ins/reference/manifest/iconurl
[defaultsettings (contentapp)]: /office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: /office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: https://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissions (contentapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (taskpaneapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (mailapp)]: /office/dev/add-ins/reference/manifest/permissions
[rule (rulecollection)]: /office/dev/add-ins/reference/manifest/rule
[rule (mailapp)]: /office/dev/add-ins/reference/manifest/rule
[requirements (mailapp)*]: /office/dev/add-ins/reference/manifest/requirements
[set*]: /office/dev/add-ins/reference/manifest/set
[sets (mailapprequirements)*]: /office/dev/add-ins/reference/manifest/sets
[form*]: /office/dev/add-ins/reference/manifest/form
[formsettings*]: /office/dev/add-ins/reference/manifest/formsettings
[sets (requirements)*]: /office/dev/add-ins/reference/manifest/sets
[hosts*]: /office/dev/add-ins/reference/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="776d6-216">ホストするための要件</span><span class="sxs-lookup"><span data-stu-id="776d6-216">Hosting requirements</span></span>

<span data-ttu-id="776d6-217">[アドイン コマンド][]などで使用されるすべてのイメージ URI はキャッシュをサポートしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="776d6-217">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="776d6-218">イメージをホストしているサーバーは、HTTP 応答で `no-cache`、`no-store`、または同様のオプションを指定する `Cache-Control` ヘッダーを返しません。</span><span class="sxs-lookup"><span data-stu-id="776d6-218">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="776d6-219">[SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 要素で指定されるソース ファイルの場所など、すべての URL は **SSL (HTTPS) でセキュリティ保護されている**べきです。</span><span class="sxs-lookup"><span data-stu-id="776d6-219">All URLs, such as the source file locations specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="776d6-220">AppSource に提出するためのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="776d6-220">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="776d6-p104">アドイン ID が有効で、一意の GUID であることを確認してください。Web 上で、一意の GUID を作成するために使用できるさまざまな GUID ジェネレーター ツールを利用できます。</span><span class="sxs-lookup"><span data-stu-id="776d6-p104">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="776d6-223">AppSource に提出するアドインには、[SupportUrl](/office/dev/add-ins/reference/manifest/supporturl) 要素も含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="776d6-223">Add-ins submitted to AppSource must also include the [SupportUrl](/office/dev/add-ins/reference/manifest/supporturl) element.</span></span> <span data-ttu-id="776d6-224">詳細については、「[AppSource に提出されたアプリとアドインの検証ポリシー](/office/dev/store/validation-policies)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="776d6-224">For more information, see [Validation policies for apps and add-ins submitted to AppSource](/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="776d6-225">必ず [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) 要素を使い、認証シナリオのために [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 要素で指定されたもの以外のドメインを指定してください。</span><span class="sxs-lookup"><span data-stu-id="776d6-225">Only use the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="776d6-226">アドイン ウィンドウで開くドメインの指定</span><span class="sxs-lookup"><span data-stu-id="776d6-226">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="776d6-227">Office on the web で実行している場合、作業ウィンドウは任意の URL に移動できます。</span><span class="sxs-lookup"><span data-stu-id="776d6-227">When running in Office Online, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="776d6-228">ただし、デスクトップ プラットフォームでは、アドインがスタート ページ (マニフェスト ファイルの [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 要素で指定されるページ) をホストするドメインとは異なるドメインの URL に移動しようとすると、移動先の URL は Office ホスト アプリケーションのアドイン ウィンドウとは別の新しいブラウザー ウィンドウで開かれます。</span><span class="sxs-lookup"><span data-stu-id="776d6-228">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="776d6-229">このデスクトップの Office の動作を変更するには、マニフェスト ファイルの [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) 要素で指定するドメインの一覧で、アドイン ウィンドウで開く各ドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="776d6-229">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="776d6-230">アドインがこの一覧にあるドメインの URL に移動しようとすると、Office on the web とデスクトップの Office の両方の作業ウィンドウで開きます。</span><span class="sxs-lookup"><span data-stu-id="776d6-230">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online.</span></span> <span data-ttu-id="776d6-231">この一覧にない URL に移動しようとすると、その URL はデスクトップの Office 新しいブラウザー ウィンドウ (アドイン ウィンドウとは別のウィンドウ) で開きます。</span><span class="sxs-lookup"><span data-stu-id="776d6-231">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="776d6-232">この動作に対する例外は 2 つあります。</span><span class="sxs-lookup"><span data-stu-id="776d6-232">There are implications to this behavior:</span></span>
> 
> - <span data-ttu-id="776d6-233">これは、アドインのルート ウィンドウに対してのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="776d6-233">This behavior applies only to the root pane of the add-in.</span></span> <span data-ttu-id="776d6-234">アドインページに iframe が埋め込まれている場合、Office デスクトップの場合でも、**AppDomains** の一覧にあるかどうかにかかわらず、その iframe を任意の URL に転送できます。</span><span class="sxs-lookup"><span data-stu-id="776d6-234">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>
> - <span data-ttu-id="776d6-235">[displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API でダイアログを開く場合、メソッドに渡される URL はアドインと同じドメインにある必要がありますが、ダイアログはデスクトップ Office であっても **AppDomains** にリストされているかどうかに関係なく、任意の URL にリダイレクトできます。</span><span class="sxs-lookup"><span data-stu-id="776d6-235">When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span> 

<span data-ttu-id="776d6-236">次に示す XML マニフェストの例では、**SourceLocation** 要素に指定された `https://www.contoso.com` ドメインでメイン アドイン ページをホストします。</span><span class="sxs-lookup"><span data-stu-id="776d6-236">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="776d6-237">また、この例では、**AppDomains** 要素リスト内の [AppDomain](/office/dev/add-ins/reference/manifest/appdomain) 要素の `https://www.northwindtraders.com` ドメインも指定しています。</span><span class="sxs-lookup"><span data-stu-id="776d6-237">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](/office/dev/add-ins/reference/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="776d6-238">アドインが www.northwindtraders.com ドメイン内のページに移動すると、Office デスクトップの場合でも、そのページはアドイン ウィンドウで開きます。</span><span class="sxs-lookup"><span data-stu-id="776d6-238">If the add-in goes to a page in the www.northwindtraders.com domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a><span data-ttu-id="776d6-239">Office.js API 呼び出しが行われるドメインを指定する</span><span class="sxs-lookup"><span data-stu-id="776d6-239">Specify domains from which Office.js API calls are made</span></span>

<span data-ttu-id="776d6-240">アドインは、マニフェスト ファイルの [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 要素で参照されているドメインから Office.js API 呼び出しを行うことができます。</span><span class="sxs-lookup"><span data-stu-id="776d6-240">Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file.</span></span> <span data-ttu-id="776d6-241">アドイン内に、Office.js API にアクセスする必要がある他の IFrame がある場合は、マニフェスト ファイルの [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) 要素で指定されているリストにそのソース URL のドメインを追加します。</span><span class="sxs-lookup"><span data-stu-id="776d6-241">If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="776d6-242">`AppDomains` リストに含まれていないソースを持つ IFrame が Office.js API 呼び出しを行おうとすると、アドインには[アクセス許可の拒否エラー](../reference/javascript-api-for-office-error-codes.md)が返されます。</span><span class="sxs-lookup"><span data-stu-id="776d6-242">If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).</span></span> 

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="776d6-243">マニフェスト v1.1 XML ファイルの例とスキーマ</span><span class="sxs-lookup"><span data-stu-id="776d6-243">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="776d6-244">後続の各セクションでは、コンテンツ アドイン、作業ウィンドウ アドイン、および Outlook アドインのマニフェスト v1.1 XML ファイルの例を示します。</span><span class="sxs-lookup"><span data-stu-id="776d6-244">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="776d6-245">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="776d6-245">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="776d6-246">作業ウィンドウ アプリ マニフェストのスキーマ</span><span class="sxs-lookup"><span data-stu-id="776d6-246">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />

  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="776d6-247">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="776d6-247">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="776d6-248">コンテンツ アプリ マニフェストのスキーマ</span><span class="sxs-lookup"><span data-stu-id="776d6-248">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />

  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="776d6-249">メール</span><span class="sxs-lookup"><span data-stu-id="776d6-249">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="776d6-250">メール アプリ マニフェストのスキーマ</span><span class="sxs-lookup"><span data-stu-id="776d6-250">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="776d6-251">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="776d6-251">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="776d6-p111">マニフェストの問題をトラブルシューティングするには、「[マニフェストの問題を検証し、トラブルシューティングする](../testing/troubleshoot-manifest.md)」を参照してください。[XML スキーマ定義 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) に対してマニフェストを検証する方法、およびランタイムのログを使用してマニフェストをデバッグする方法についての情報を確認できます。</span><span class="sxs-lookup"><span data-stu-id="776d6-p111">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="776d6-254">関連項目</span><span class="sxs-lookup"><span data-stu-id="776d6-254">See also</span></span>

* <span data-ttu-id="776d6-255">[マニフェストでアドイン コマンドを作成する]、[アドイン コマンド]</span><span class="sxs-lookup"><span data-stu-id="776d6-255">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="776d6-256">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="776d6-256">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="776d6-257">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="776d6-257">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="776d6-258">Office アドイン マニフェストのスキーマ参照</span><span class="sxs-lookup"><span data-stu-id="776d6-258">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="776d6-259">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="776d6-259">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[アドイン コマンド]: create-addin-commands.md
[add-in commands]: create-addin-commands.md
