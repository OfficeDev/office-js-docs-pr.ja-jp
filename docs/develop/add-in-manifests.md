---
title: Office アドインの XML マニフェスト
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: 71c77e190d5d2d6cc67ada671b9efe3168b7f7b5
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703820"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="d2b16-102">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="d2b16-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="d2b16-103">Office アドインの XML マニフェスト ファイルでは、エンド ユーザーが Office ドキュメントや Office アプリケーションにアドインをインストールして使用するときにアドインをアクティブ化する方法が記述されています。</span><span class="sxs-lookup"><span data-stu-id="d2b16-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="d2b16-104">このスキーマに基づいた XML マニフェスト ファイルを使用すると、Office アドインで次のことができます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="d2b16-105">ID、バージョン、説明、表示名、および既定のロケールを指定することで、アプリ自体について説明する。</span><span class="sxs-lookup"><span data-stu-id="d2b16-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="d2b16-106">アドインのブランド化に使用するイメージと、Office リボンで[アドイン コマンド][]に使用する画像を指定する。</span><span class="sxs-lookup"><span data-stu-id="d2b16-106">Specify the images used for branding the Add-in and iconography used for [Add-in Commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="d2b16-107">アドインを Office に統合する方法を指定する。アドインによって作成されるカスタム UI (リボンのボタンなど) の統合も含む。</span><span class="sxs-lookup"><span data-stu-id="d2b16-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="d2b16-108">コンテンツ アドインに必要な既定のサイズ、および Outlook アドインに必要な高さを指定する。</span><span class="sxs-lookup"><span data-stu-id="d2b16-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="d2b16-109">ドキュメントの読み取り、書き込みなど、Office アドインに必要なアクセス許可を宣言する。</span><span class="sxs-lookup"><span data-stu-id="d2b16-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="d2b16-110">Outlook アドインでは、アプリがアクティブ化されてメッセージ、予定、または会議出席依頼アイテムを操作するコンテキストを指定するルールを定義する。</span><span class="sxs-lookup"><span data-stu-id="d2b16-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="d2b16-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="d2b16-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="d2b16-113">必要な要素</span><span class="sxs-lookup"><span data-stu-id="d2b16-113">Required elements</span></span>

<span data-ttu-id="d2b16-114">次の表に、3 種類の Office アドインに必要な要素を示します。</span><span class="sxs-lookup"><span data-stu-id="d2b16-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="d2b16-115">Office アドインの種類ごとの必要な要素</span><span class="sxs-lookup"><span data-stu-id="d2b16-115">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="d2b16-116">要素</span><span class="sxs-lookup"><span data-stu-id="d2b16-116">Element</span></span>                                                                                      | <span data-ttu-id="d2b16-117">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="d2b16-117">Content</span></span> | <span data-ttu-id="d2b16-118">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="d2b16-118">Task pane</span></span> | <span data-ttu-id="d2b16-119">Outlook</span><span class="sxs-lookup"><span data-stu-id="d2b16-119">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="d2b16-120">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-120">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="d2b16-121">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-121">X</span></span>    |     <span data-ttu-id="d2b16-122">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-122">X</span></span>     |    <span data-ttu-id="d2b16-123">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-123">X</span></span>    |
| <span data-ttu-id="d2b16-124">[Id][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-124">[Id][]</span></span>                                                                                       |    <span data-ttu-id="d2b16-125">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-125">X</span></span>    |     <span data-ttu-id="d2b16-126">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-126">X</span></span>     |    <span data-ttu-id="d2b16-127">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-127">X</span></span>    |
| <span data-ttu-id="d2b16-128">[Version][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-128">[Version][]</span></span>                                                                                  |    <span data-ttu-id="d2b16-129">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-129">X</span></span>    |     <span data-ttu-id="d2b16-130">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-130">X</span></span>     |    <span data-ttu-id="d2b16-131">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-131">X</span></span>    |
| <span data-ttu-id="d2b16-132">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-132">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="d2b16-133">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-133">X</span></span>    |     <span data-ttu-id="d2b16-134">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-134">X</span></span>     |    <span data-ttu-id="d2b16-135">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-135">X</span></span>    |
| <span data-ttu-id="d2b16-136">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-136">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="d2b16-137">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-137">X</span></span>    |     <span data-ttu-id="d2b16-138">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-138">X</span></span>     |    <span data-ttu-id="d2b16-139">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-139">X</span></span>    |
| <span data-ttu-id="d2b16-140">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-140">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="d2b16-141">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-141">X</span></span>    |     <span data-ttu-id="d2b16-142">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-142">X</span></span>     |    <span data-ttu-id="d2b16-143">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-143">X</span></span>    |
| <span data-ttu-id="d2b16-144">[説明][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-144">[Description][]</span></span>                                                                              |    <span data-ttu-id="d2b16-145">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-145">X</span></span>    |     <span data-ttu-id="d2b16-146">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-146">X</span></span>     |    <span data-ttu-id="d2b16-147">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-147">X</span></span>    |
| <span data-ttu-id="d2b16-148">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-148">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="d2b16-149">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-149">X</span></span>    |     <span data-ttu-id="d2b16-150">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-150">X</span></span>     |    <span data-ttu-id="d2b16-151">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-151">X</span></span>    |
| <span data-ttu-id="d2b16-152">[HighResolutionIconUrl][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-152">[HighResolutionIconUrl][]</span></span>                                                                    |    <span data-ttu-id="d2b16-153">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-153">X</span></span>    |     <span data-ttu-id="d2b16-154">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-154">X</span></span>     |    <span data-ttu-id="d2b16-155">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-155">X</span></span>    |
| <span data-ttu-id="d2b16-156">[DefaultSettings（ContentApp）][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-156">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="d2b16-157">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-157">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="d2b16-158">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-158">X</span></span>    |     <span data-ttu-id="d2b16-159">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-159">X</span></span>     |         |
| <span data-ttu-id="d2b16-160">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-160">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="d2b16-161">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-161">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="d2b16-162">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-162">X</span></span>    |     <span data-ttu-id="d2b16-163">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-163">X</span></span>     |         |
| <span data-ttu-id="d2b16-164">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-164">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="d2b16-165">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-165">X</span></span>    |
| <span data-ttu-id="d2b16-166">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-166">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="d2b16-167">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-167">X</span></span>    |
| <span data-ttu-id="d2b16-168">[Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-168">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="d2b16-169">[Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-169">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="d2b16-170">[Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-170">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="d2b16-171">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-171">X</span></span>    |     <span data-ttu-id="d2b16-172">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-172">X</span></span>     |    <span data-ttu-id="d2b16-173">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-173">X</span></span>    |
| <span data-ttu-id="d2b16-174">[Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-174">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="d2b16-175">[Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-175">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="d2b16-176">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-176">X</span></span>    |
| <span data-ttu-id="d2b16-177">[Requirements (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-177">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="d2b16-178">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-178">X</span></span>    |
| <span data-ttu-id="d2b16-179">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-179">[Set\*][]</span></span><br/><span data-ttu-id="d2b16-180">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-180">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="d2b16-181">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-181">X</span></span>    |
| <span data-ttu-id="d2b16-182">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-182">[Form\*][]</span></span><br/><span data-ttu-id="d2b16-183">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-183">[**FormSettings][]</span></span>                                                              |         |           |    <span data-ttu-id="d2b16-184">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-184">X</span></span>    |
| <span data-ttu-id="d2b16-185">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-185">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="d2b16-186">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-186">X</span></span>    |     <span data-ttu-id="d2b16-187">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-187">X</span></span>     |         |
| <span data-ttu-id="d2b16-188">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="d2b16-188">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="d2b16-189">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-189">X</span></span>    |     <span data-ttu-id="d2b16-190">X</span><span class="sxs-lookup"><span data-stu-id="d2b16-190">X</span></span>     |         |

<span data-ttu-id="d2b16-191">_\*Office アドイン マニフェスト スキーマ バージョン 1.1 で追加されました。_</span><span class="sxs-lookup"><span data-stu-id="d2b16-191">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: http://msdn.microsoft.com/en-us/library/68f1cada-66f8-4341-45f5-14e0634c24fb%28Office.15%29.aspx
[id]: http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx
[version]: http://msdn.microsoft.com/en-us/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx
[providername]: http://msdn.microsoft.com/en-us/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx
[defaultlocale]: http://msdn.microsoft.com/en-us/library/04796a3a-3afa-dc85-db66-4677560c185c%28Office.15%29.aspx
[displayname]: http://msdn.microsoft.com/en-us/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx
[description]: http://msdn.microsoft.com/en-us/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx
[iconurl]: http://msdn.microsoft.com/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx
[highresolutioniconurl]: http://msdn.microsoft.com/library/ff7b2647-ec8e-70dc-4e4a-e1a1377ff3f2%28Office.15%29.aspx
[defaultsettings (contentapp)]: http://msdn.microsoft.com/en-us/library/f7edc689-551f-1a17-ea81-ffd58f534557%28Office.15%29.aspx
[defaultsettings (taskpaneapp)]: http://msdn.microsoft.com/en-us/library/36e3d139-56a4-fb3d-0a21-cbd14e606765%28Office.15%29.aspx
[sourcelocation (contentapp)]: http://msdn.microsoft.com/en-us/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx
[sourcelocation (taskpaneapp)]: http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx
[desktopsettings]: http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/en-us/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissions (contentapp)]: http://msdn.microsoft.com/en-us/library/9f3dcf9c-fced-c115-4f0d-38d60fb7c583%28Office.15%29.aspx
[permissions (taskpaneapp)]: http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx
[permissions (mailapp)]: http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx
[rule (rulecollection)]: http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx
[rule (mailapp)]: http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx
[requirements (mailapp)*]: http://msdn.microsoft.com/en-us/library/9536ea30-34f7-76b5-7f30-1508626840e4%28Office.15%29.aspx
[set*]: http://msdn.microsoft.com/en-us/library/1506daa1-332c-30e1-6402-3371bcd0b895%28Office.15%29.aspx
[sets (mailapprequirements)*]: http://msdn.microsoft.com/en-us/library/2a6a2484-eeee-37e4-43bc-c185e8ae0d1d%28Office.15%29.aspx
[form*]: http://msdn.microsoft.com/en-us/library/77a8ac83-c22b-1225-4fc4-ba4038b68648%28Office.15%29.aspx
[formsettings*]: http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx
[sets (requirements)*]: http://msdn.microsoft.com/en-us/library/509be287-b532-87c6-71ac-64f3a4bbd3af%28Office.15%29.aspx
[hosts*]: http://msdn.microsoft.com/library/f9a739c1-3daf-c03a-2bd9-4a2a6b870101%28Office.15%29.aspx

## <a name="hosting-requirements"></a><span data-ttu-id="d2b16-219">ホストするための要件</span><span class="sxs-lookup"><span data-stu-id="d2b16-219">Hosting requirements</span></span>

<span data-ttu-id="d2b16-220">[アドイン コマンド][]などで使用されるすべてのイメージ URI はキャッシュをサポートしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="d2b16-220">All image URIs, such as those used for [Add-in Commands][], must support caching.</span></span> <span data-ttu-id="d2b16-221">イメージをホストしているサーバーは、HTTP 応答で `no-cache`、`no-store`、または同様のオプションを指定する `Cache-Control` ヘッダーを返しません。</span><span class="sxs-lookup"><span data-stu-id="d2b16-221">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="d2b16-222">[SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) 要素で指定されるソース ファイルの場所など、すべての URL は **SSL (HTTPS) でセキュリティ保護されている**べきです。</span><span class="sxs-lookup"><span data-stu-id="d2b16-222">All URLs, such as the source file locations specified in the [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="d2b16-223">AppSource に提出するためのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="d2b16-223">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="d2b16-p103">アドイン ID が有効で、一意の GUID であることを確認してください。Web 上で、一意の GUID を作成するために使用できるさまざまな GUID ジェネレーター ツールを利用できます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="d2b16-226">AppSource に提出するアドインには、[SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl) 要素も含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="d2b16-226">Add-ins submitted to AppSource must also include the [SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl) element.</span></span> <span data-ttu-id="d2b16-227">詳細については、「[AppSource に提出されたアプリとアドインの検証ポリシー](https://docs.microsoft.com/office/dev/store/validation-policies)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d2b16-227">For more information, see [Validation policies for apps and add-ins submitted to AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="d2b16-228">必ず [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) 要素を使い、認証シナリオのために [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) 要素で指定されたもの以外のドメインを指定してください。</span><span class="sxs-lookup"><span data-stu-id="d2b16-228">Only use the [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="d2b16-229">アドイン ウィンドウで開くドメインの指定</span><span class="sxs-lookup"><span data-stu-id="d2b16-229">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="d2b16-230">Office Onlineで実行している場合、作業ウィンドウは任意の URL にナビゲートできます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-230">When running in Office Online, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="d2b16-231">ただし、デスクトップ プラットフォームでは、アドインがスタート ページ (アドインのマニフェスト ファイルの [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) 要素で指定されるページ) をホストするドメインとは異なるドメインの URL に移動しようとすると、移動先の URL は Office ホスト アプリケーションのアドイン ウィンドウとは別の新しいブラウザー ウィンドウで開かれます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-231">By default, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) element of the manifest file), that URL will open in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="d2b16-232">この動作 (Office デスクトップ) を変更するには、マニフェスト ファイルの [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) 要素で指定するドメインの一覧で、アドイン ウィンドウで開く各ドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="d2b16-232">To override this behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="d2b16-233">この一覧にないドメインの URL にアドインが移動しようとすると、その URL が Office デスクトップと Office Online の両方の作業ウィンドウで開きます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-233">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online.</span></span> <span data-ttu-id="d2b16-234">この一覧にない URL に移動しようとすると、Office デスクトップでは、その URL が新しいブラウザー ウィンドウ (アドイン ウィンドウとは別のウィンドウ) で開きます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-234">If the add-in tries to go to a URL in a domain that isn't in the list, that URL will open in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="d2b16-235">この現象は、アドインのルート ウィンドウにのみ当てはまります。</span><span class="sxs-lookup"><span data-stu-id="d2b16-235">This behavior applies only to the root pane of the add-in.</span></span> <span data-ttu-id="d2b16-236">アドインページに iframe が埋め込まれている場合、Office デスクトップの場合でも、**AppDomains** の一覧にあるかどうかにかかわらず、その iframe を任意の URL にダイレクトすることができます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-236">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="d2b16-237">次に示す XML マニフェストの例では、**SourceLocation** 要素に指定された `https://www.contoso.com` ドメインでメイン アドイン ページをホストします。</span><span class="sxs-lookup"><span data-stu-id="d2b16-237">The following XML manifest example hosts its main add-in page in the  `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="d2b16-238">また、この例では、**AppDomains** 要素リストの [AppDomain](https://dev.office.com/reference/add-ins/manifest/appdomain) 要素で `https://www.northwindtraders.com` ドメインも指定しています。</span><span class="sxs-lookup"><span data-stu-id="d2b16-238">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](https://dev.office.com/reference/add-ins/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="d2b16-239">アドインが www.northwindtraders.com ドメイン内のページに移動すると、Office デスクトップでも、そのページがアドイン ウィンドウで開きます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-239">If the add-in goes to a page in the www.northwindtraders.com domain, that page will open in the add-in pane.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
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

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="d2b16-240">マニフェスト v1.1 XML ファイルの例とスキーマ</span><span class="sxs-lookup"><span data-stu-id="d2b16-240">Manifest v1.1 XML file examples and schemas</span></span>
<span data-ttu-id="d2b16-241">後続の各セクションでは、コンテンツ アドイン、作業ウィンドウ アドイン、および Outlook アドインのマニフェスト v1.1 XML ファイルの例を示します。</span><span class="sxs-lookup"><span data-stu-id="d2b16-241">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="d2b16-242">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="d2b16-242">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="d2b16-243">作業ウィンドウ アプリ マニフェストのスキーマ</span><span class="sxs-lookup"><span data-stu-id="d2b16-243">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

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
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

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
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
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
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
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
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
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

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="d2b16-244">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="d2b16-244">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="d2b16-245">コンテンツ アプリ マニフェストのスキーマ</span><span class="sxs-lookup"><span data-stu-id="d2b16-245">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

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
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
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

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="d2b16-246">メール</span><span class="sxs-lookup"><span data-stu-id="d2b16-246">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="d2b16-247">メール アプリ マニフェストのスキーマ</span><span class="sxs-lookup"><span data-stu-id="d2b16-247">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

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
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="d2b16-248">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="d2b16-248">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="d2b16-p109">マニフェストの問題をトラブルシューティングするには、「[マニフェストの問題を検証し、トラブルシューティングする](../testing/troubleshoot-manifest.md)」を参照してください。[XML スキーマ定義 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) に対してマニフェストを検証する方法、およびランタイムのログを使用してマニフェストをデバッグする方法についての情報を確認できます。</span><span class="sxs-lookup"><span data-stu-id="d2b16-p109">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="d2b16-251">関連項目</span><span class="sxs-lookup"><span data-stu-id="d2b16-251">See also</span></span>

* <span data-ttu-id="d2b16-252">[マニフェストでアドイン コマンドを作成する]、[アドイン コマンド]</span><span class="sxs-lookup"><span data-stu-id="d2b16-252">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="d2b16-253">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="d2b16-253">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="d2b16-254">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="d2b16-254">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="d2b16-255">Office アドイン マニフェストのスキーマ参照</span><span class="sxs-lookup"><span data-stu-id="d2b16-255">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="d2b16-256">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="d2b16-256">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[アドイン コマンド]: create-addin-commands.md
[add-in commands]: create-addin-commands.md