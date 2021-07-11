---
title: Outlook アドインにモバイル サポートを追加する
description: Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: f653f43228c7667bc6848d4f0a6d2e9fd1768964
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349008"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a><span data-ttu-id="ca65c-103">Outlook Mobile のアドイン コマンドのサポートを追加する</span><span class="sxs-lookup"><span data-stu-id="ca65c-103">Add support for add-in commands for Outlook Mobile</span></span>

<span data-ttu-id="ca65c-104">Outlook Mobile でアドイン コマンドを使用すると、ユーザーは Outlook on the web、Windows、および Mac で既に持っているのと同じ機能 (いくつかの制限[があります)](#code-considerations)にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="ca65c-104">Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="ca65c-105">Outlook Mobile のサポートを追加するには、アドイン マニフェストを更新する必要があります。さらに、モバイル シナリオのコードを変更することが必要な場合もあります。</span><span class="sxs-lookup"><span data-stu-id="ca65c-105">Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.</span></span>

## <a name="updating-the-manifest"></a><span data-ttu-id="ca65c-106">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="ca65c-106">Updating the manifest</span></span>

<span data-ttu-id="ca65c-p102">Outlook Mobile でアドイン コマンドを有効にするための最初の手順は、アドイン マニフェストでの定義です。[VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 スキーマは、モバイル用に新しいフォーム ファクター [MobileFormFactor](../reference/manifest/mobileformfactor.md) を定義します。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p102">The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span></span>

<span data-ttu-id="ca65c-p103">この要素には、モバイル クライアントにアドインを読み込むためのすべての情報が含まれています。これにより、モバイル エクスペリエンスに対して完全に異なる UI 要素と JavaScript ファイルを定義することができます。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p103">This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.</span></span>

<span data-ttu-id="ca65c-111">次の使用例は、要素内の 1 つの作業ウィンドウ ボタンを示 `MobileFormFactor` しています。</span><span class="sxs-lookup"><span data-stu-id="ca65c-111">The following example shows a single task pane button in a `MobileFormFactor` element.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

<span data-ttu-id="ca65c-112">これは、[DesktopFormFactor](../reference/manifest/desktopformfactor.md) 要素に表示される要素と非常によく似ていますが、いくつかの注目すべき違いがあります。</span><span class="sxs-lookup"><span data-stu-id="ca65c-112">This is very similar to the elements that appear in a [DesktopFormFactor](../reference/manifest/desktopformfactor.md) element, with some notable differences.</span></span>

- <span data-ttu-id="ca65c-113">[OfficeTab](../reference/manifest/officetab.md) 要素は使用されません。</span><span class="sxs-lookup"><span data-stu-id="ca65c-113">The [OfficeTab](../reference/manifest/officetab.md) element is not used.</span></span>
- <span data-ttu-id="ca65c-p104">[ExtensionPoint](../reference/manifest/extensionpoint.md) 要素に含まれる子要素は 1 つでなければなりません。アドインがボタンを 1 つのみ追加する場合、子要素は [Control](../reference/manifest/control.md) 要素になります。アドインがボタンを複数追加する場合、子要素は複数の `Control` 要素を含む [Group](../reference/manifest/group.md) 要素になります。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p104">The [ExtensionPoint](../reference/manifest/extensionpoint.md) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](../reference/manifest/control.md) element. If the add-in adds more than one button, the child element should be a [Group](../reference/manifest/group.md) element that contains multiple `Control` elements.</span></span>
- <span data-ttu-id="ca65c-117">`Control` 要素に相当する `Menu` の種類はありません。</span><span class="sxs-lookup"><span data-stu-id="ca65c-117">There is no `Menu` type equivalent for the `Control` element.</span></span>
- <span data-ttu-id="ca65c-118">[Supertip](../reference/manifest/supertip.md) 要素は使用されません。</span><span class="sxs-lookup"><span data-stu-id="ca65c-118">The [Supertip](../reference/manifest/supertip.md) element is not used.</span></span>
- <span data-ttu-id="ca65c-p105">アイコンの必須サイズが異なります。モバイル アドインは少なくとも 25x25、32x32 および 48x48 ピクセルのアイコンをサポートする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p105">The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.</span></span>

## <a name="code-considerations"></a><span data-ttu-id="ca65c-121">コードに関する考慮事項</span><span class="sxs-lookup"><span data-stu-id="ca65c-121">Code considerations</span></span>

<span data-ttu-id="ca65c-122">モバイル用のアドインの設計には、追加の考慮事項がいくつか導入されています。</span><span class="sxs-lookup"><span data-stu-id="ca65c-122">Designing an add-in for mobile introduces some additional considerations.</span></span>

### <a name="use-rest-instead-of-exchange-web-services"></a><span data-ttu-id="ca65c-123">Exchange Web サービスの代わりに REST を使用する</span><span class="sxs-lookup"><span data-stu-id="ca65c-123">Use REST instead of Exchange Web Services</span></span>

<span data-ttu-id="ca65c-p106">[Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドは、Outlook Mobile ではサポートされていません。可能な場合には、アドインは優先的に Office.js API から情報を取得します。Office.js API によって表示されていない情報がアドインで必要な場合、[Outlook REST APIs](/outlook/rest/) を使用してユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p106">The [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.</span></span>

<span data-ttu-id="ca65c-127">メールボックス要件セット 1.5 では、REST API と互換性のあるアクセス トークンを要求できる新しいバージョンの[Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)と、ユーザーの REST API エンドポイントの検索に使用できる新しい[Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)プロパティが導入されました。</span><span class="sxs-lookup"><span data-stu-id="ca65c-127">Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property that can be used to find the REST API endpoint for the user.</span></span>

### <a name="pinch-zoom"></a><span data-ttu-id="ca65c-128">ピンチによるズーム</span><span class="sxs-lookup"><span data-stu-id="ca65c-128">Pinch zoom</span></span>

<span data-ttu-id="ca65c-p107">既定で、ユーザーは "ピンチによるズーム" ジェスチャを使用して作業ウィンドウで拡大することができます。ご使用のシナリオでこれが該当しない場合は、HTML でピンチによるズームを無効にしてください。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p107">By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.</span></span>

### <a name="close-task-panes"></a><span data-ttu-id="ca65c-131">作業ウィンドウを閉じる</span><span class="sxs-lookup"><span data-stu-id="ca65c-131">Close task panes</span></span>

<span data-ttu-id="ca65c-p108">Outlook Mobile では、作業ウィンドウが画面全体を占めるので、既定ではユーザーが作業ウィンドウを閉じてメッセージに戻る必要があります。シナリオが完成したら、[Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) メソッドを使用して作業ウィンドウを閉じることを検討してください。</span><span class="sxs-lookup"><span data-stu-id="ca65c-p108">In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) method to close the task pane when your scenario is complete.</span></span>

### <a name="compose-mode-and-appointments"></a><span data-ttu-id="ca65c-134">作成モードと予定</span><span class="sxs-lookup"><span data-stu-id="ca65c-134">Compose mode and appointments</span></span>

<span data-ttu-id="ca65c-135">現在、Outlook Mobile のアドインは、メッセージ読み取り時のアクティブ化のみをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="ca65c-135">Currently add-ins in Outlook Mobile only support activation when reading messages.</span></span> <span data-ttu-id="ca65c-136">メッセージを作成するときや、予定を表示または作成するときには、アドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="ca65c-136">Add-ins are not activated when composing messages or when viewing or composing appointments.</span></span> <span data-ttu-id="ca65c-137">ただし、オンライン会議プロバイダー統合アドインは、予定オーガナイザー モードでアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="ca65c-137">However, online meeting provider integrated add-ins can be activated in Appointment Organizer mode.</span></span> <span data-ttu-id="ca65c-138">この例外[の詳細についてはOutlook](online-meeting.md)会議プロバイダーのモバイル アドインの作成に関する記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca65c-138">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this exception.</span></span>

### <a name="unsupported-apis"></a><span data-ttu-id="ca65c-139">サポートされていない API</span><span class="sxs-lookup"><span data-stu-id="ca65c-139">Unsupported APIs</span></span>

<span data-ttu-id="ca65c-140">要件セット 1.6 以降で導入された API は、モバイル ではOutlookされません。</span><span class="sxs-lookup"><span data-stu-id="ca65c-140">APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile.</span></span> <span data-ttu-id="ca65c-141">以前の要件セットの次の API もサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ca65c-141">The following APIs from earlier requirement sets are also not supported.</span></span>

- [<span data-ttu-id="ca65c-142">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="ca65c-142">Office.context.officeTheme</span></span>](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
- [<span data-ttu-id="ca65c-143">Office.context.mailbox.ewsUrl</span><span class="sxs-lookup"><span data-stu-id="ca65c-143">Office.context.mailbox.ewsUrl</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
- [<span data-ttu-id="ca65c-144">Office.context.mailbox.convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="ca65c-144">Office.context.mailbox.convertToEwsId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="ca65c-145">Office.context.mailbox.convertToRestId</span><span class="sxs-lookup"><span data-stu-id="ca65c-145">Office.context.mailbox.convertToRestId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="ca65c-146">Office.context.mailbox.displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ca65c-146">Office.context.mailbox.displayAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="ca65c-147">Office.context.mailbox.displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="ca65c-147">Office.context.mailbox.displayMessageForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="ca65c-148">Office.context.mailbox.displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ca65c-148">Office.context.mailbox.displayNewAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="ca65c-149">Office.context.mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="ca65c-149">Office.context.mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="ca65c-150">Office.context.mailbox.item.dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ca65c-150">Office.context.mailbox.item.dateTimeModified</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
- [<span data-ttu-id="ca65c-151">Office.context.mailbox.item.displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ca65c-151">Office.context.mailbox.item.displayReplyAllForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="ca65c-152">Office.context.mailbox.item.displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ca65c-152">Office.context.mailbox.item.displayReplyForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="ca65c-153">Office.context.mailbox.item.getEntities</span><span class="sxs-lookup"><span data-stu-id="ca65c-153">Office.context.mailbox.item.getEntities</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="ca65c-154">Office.context.mailbox.item.getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ca65c-154">Office.context.mailbox.item.getEntitiesByType</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="ca65c-155">Office.context.mailbox.item.getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ca65c-155">Office.context.mailbox.item.getFilteredEntitiesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="ca65c-156">Office.context.mailbox.item.getRegexMatches</span><span class="sxs-lookup"><span data-stu-id="ca65c-156">Office.context.mailbox.item.getRegexMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="ca65c-157">Office.context.mailbox.item.getRegexMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ca65c-157">Office.context.mailbox.item.getRegexMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a><span data-ttu-id="ca65c-158">関連項目</span><span class="sxs-lookup"><span data-stu-id="ca65c-158">See also</span></span>

[<span data-ttu-id="ca65c-159">Exchange サーバーと Outlook クライアントでサポートされる要件セット</span><span class="sxs-lookup"><span data-stu-id="ca65c-159">Requirement sets supported by Exchange servers and Outlook clients</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)