---
title: マニフェスト要素の正しい順序を確認する方法
description: 親要素内で子要素を配置するための正しい順序を確認する方法について説明します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 2ee80167a76861209e814dc6c272720feb3a9cf1
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173914"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="65d6f-103">マニフェスト要素の正しい順序を確認する方法</span><span class="sxs-lookup"><span data-stu-id="65d6f-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="65d6f-104">Office アドインのマニフェストの XML 要素は適切な親要素の下に配置する必要があり、*また*、親要素の下で子要素同士が特定の順序に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="65d6f-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="65d6f-105">必要な順序は、[[スキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)] フォルダー内の XSD ファイルで指定されています。</span><span class="sxs-lookup"><span data-stu-id="65d6f-105">The required ordering is specified in the XSD files in the [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) folder.</span></span> <span data-ttu-id="65d6f-106">XSD ファイルは、作業ウィンドウ、コンテンツ、およびメール アドインのサブフォルダーに分類されます。</span><span class="sxs-lookup"><span data-stu-id="65d6f-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="65d6f-107">例えば、`<OfficeApp>` 要素では、`<Id>`、`<Version>`、`<ProviderName>` はこの順序で表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="65d6f-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="65d6f-108">`<AlternateId>` 要素が追加された場合、この要素は `<Id>` 要素と `<Version>` 要素の間に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="65d6f-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="65d6f-109">順序が間違っている要素が 1 つでもあると、マニフェストは有効にならず、アドインも読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="65d6f-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="65d6f-110">[office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest)内の検証機能は、要素が正しい親の下にある場合と同じエラー メッセージを使用します。</span><span class="sxs-lookup"><span data-stu-id="65d6f-110">The [validator within office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="65d6f-111">エラーには、子要素が親要素の有効な子ではないと表示されます。</span><span class="sxs-lookup"><span data-stu-id="65d6f-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="65d6f-112">そのようなエラーが表示されるものの、子要素のレファレンス ドキュメントがこの子要素は親要素の有効な子 *である* と示す場合は、おそらく、子要素が間違った順序で配置されていることが原因です。</span><span class="sxs-lookup"><span data-stu-id="65d6f-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="65d6f-113">次のセクションでは、マニフェスト要素を表示する順序で示します。</span><span class="sxs-lookup"><span data-stu-id="65d6f-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="65d6f-114">要素の属性が 、 `type` `<OfficeApp>` `TaskPaneApp` `ContentApp` `MailApp` .</span><span class="sxs-lookup"><span data-stu-id="65d6f-114">There are differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="65d6f-115">これらのセクションが扱いすぎずになじむのを強くするために、非常に複雑な要素は別の `<VersionOverrides>` セクションに分かれています。</span><span class="sxs-lookup"><span data-stu-id="65d6f-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="65d6f-116">表示される要素の一部が必須ではありません。</span><span class="sxs-lookup"><span data-stu-id="65d6f-116">Not all of the elements shown are mandatory.</span></span> <span data-ttu-id="65d6f-117">スキーマ内 `minOccurs` の要素の値が **0** [](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3)の場合、要素は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="65d6f-117">If the `minOccurs` value for a element is **0** in the [schema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="65d6f-118">基本的な作業ウィンドウ アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-118">Basic task pane add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="TaskPaneApp">
    <Id>
    <AlternateID>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
        <Sets>
            <Set>
        <Methods>
            <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <Permissions>
    <Dictionary>
        <TargetDialects>
        <QueryUri>
        <CitationText>
        <DictionaryName>
        <DictionaryHomePage>
    <VersionOverrides>*
    <ExtendedOverrides>
```

<span data-ttu-id="65d6f-119">\*VersionOverrides の子要素の順序については [、VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) 内での作業ウィンドウ アドイン要素の順序を参照してください。</span><span class="sxs-lookup"><span data-stu-id="65d6f-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="65d6f-120">基本的なメール アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-120">Basic mail add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="MailApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl>
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <FormSettings>
        <Form>
        <DesktopSettings>
            <SourceLocation>
            <RequestedHeight>
        <TabletSettings>
            <SourceLocation>
            <RequestedHeight>
        <PhoneSettings>
            <SourceLocation>
    <Permissions>
    <Rule>
    <DisableEntityHighlighting>
    <VersionOverrides>*
```

<span data-ttu-id="65d6f-121">\*VersionOverrides の子要素の順序については [、VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) 内でのメール アドイン要素の順序付けと [VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) 内でのメール アドイン要素の順序付けをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="65d6f-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="65d6f-122">基本的なコンテンツ アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-122">Basic content add-in element ordering</span></span>

```xml
<OfficeApp xsi:type="ContentApp">
    <Id>
    <AlternateId>
    <Version>
    <ProviderName>
    <DefaultLocale>
    <DisplayName>
        <Override>
    <Description>
        <Override>
    <IconUrl >
        <Override>
    <HighResolutionIconUrl>
        <Override>
    <SupportUrl>
    <AppDomains>
        <AppDomain>
    <Hosts>
        <Host>
    <Requirements>
    <Sets>
        <Set>
    <Methods>
        <Method>
    <DefaultSettings>
        <SourceLocation>
            <Override>
    <RequestedWidth>
    <RequestedHeight>
    <Permissions>
    <AllowSnapshot>
    <VersionOverrides>*
```

<span data-ttu-id="65d6f-123">\*VersionOverrides [の子要素の順序については、VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) 内でのコンテンツ アドイン要素の順序を参照してください。</span><span class="sxs-lookup"><span data-stu-id="65d6f-123">\*See [Content add-in element ordering within VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="65d6f-124">VersionOverrides 内での作業ウィンドウ アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-124">Task pane add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <Runtimes>
                <Runtime>
            <AllFormFactors>
                <ExtensionPoint>
                    <Script>
                        <SourceLocation>
                    <Page>
                        <SourceLocation>
                    <Metadata>
                        <SourceLocation>
                    <Namespace>
            <DesktopFormFactor>
                <GetStarted>
                    <Title>
                    <Description>
                    <LearnMoreUrl>
                <FunctionFile>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                    <CustomTab>
                        <OverriddenByRibbonApi>
                        <Group> (can be below <ControlGroup>)
                            <OverriddenByRibbonApi>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <OverriddenByRibbonApi>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>  
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                <Enabled>
                                <Items>
                                    <Item>
                                        <OverriddenByRibbonApi>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <ControlGroup> (can be above <Group>)
                        <Label>
                        <InsertAfter> (or <InsertBefore>)
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Enabled>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
        <Resources>
            <Images>
                <Image>
                    <Override>
            <Urls>
                <Url>
                    <Override>
            <ShortStrings>
                <String>
                    <Override>
            <LongStrings>
                <String>
                    <Override>
        <WebApplicationInfo>
            <Id>
            <MsaId>
            <Resource>
            <Scopes>
                <Scope>
            <Authorizations>
                <Authorization>
                    <Resource>
                    <Scopes>
                        <Scope>
        <EquivalentAddins>
            <EquivalentAddin>
                <ProgId>
                <DisplayName>
                <FileName>
                <Type>
```

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="65d6f-125">VersionOverrides Ver 内でのメール アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-125">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="65d6f-126">1.0</span><span class="sxs-lookup"><span data-stu-id="65d6f-126">1.0</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
    <Hosts>
        <Host>
            <DesktopFormFactor>
                <ExtensionPoint>
                    <OfficeTab>
                        <Group>
                            <Label>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>
                                <Action>
                                    <SourceLocation>
                                    <FunctionName>
                    <CustomTab>
                        <Group>
                            <Label>
                            <Icon>
                                <Image>
                            <Control>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Icon>
                                    <Image>  
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                <Items>
                                    <Item>
                                        <Label>
                                        <Supertip>
                                            <Title>
                                            <Description>
                                        <Action>
                                            <TaskpaneId>
                                            <SourceLocation>
                                            <Title>
                                            <FunctionName>
                        <Label>
                    <OfficeMenu>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <VersionOverrides>*
```

<span data-ttu-id="65d6f-127">\* 値を持つ VersionOverrides は、外側の `type` `VersionOverridesV1_1` VersionOverride の末尾に入れ子にすることができます `VersionOverridesV1_0` 。</span><span class="sxs-lookup"><span data-stu-id="65d6f-127">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="65d6f-128">要素 [の順序については、VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) 内でのメール アドイン要素の順序を参照してください `VersionOverridesV1_1` 。</span><span class="sxs-lookup"><span data-stu-id="65d6f-128">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="65d6f-129">VersionOverrides Ver 内でのメール アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-129">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="65d6f-130">1.1</span><span class="sxs-lookup"><span data-stu-id="65d6f-130">1.1</span></span>

```xml
<VersionOverrides>
    <Description>
    <Requirements>
    <Sets>
        <Set>
    <Hosts>
    <Host>
        <DesktopFormFactor>
            <ExtensionPoint>
                <OfficeTab>
                    <Group>
                        <Label>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>
                            <Action>
                                <SourceLocation>
                                <FunctionName>
                <CustomTab>
                    <Group>
                        <Label>
                        <Icon>
                            <Image>
                        <Control>
                            <Label>
                            <Supertip>
                                <Title>
                                <Description>
                            <Icon>
                                <Image>  
                            <Action>
                                <TaskpaneId>
                                <SourceLocation>
                                <Title>
                                <FunctionName>
                            <Items>
                                <Item>
                                    <Label>
                                    <Supertip>
                                        <Title>
                                        <Description>
                                    <Action>
                                        <TaskpaneId>
                                        <SourceLocation>
                                        <Title>
                                        <FunctionName>
                    <Label>
                <OfficeMenu>
                    <Control>
                        <Label>
                        <Supertip>
                            <Title>
                            <Description>
                        <Icon>
                            <Image>  
                        <Action>
                            <TaskpaneId>
                            <SourceLocation>
                            <Title>
                            <FunctionName>
                        <Items>
                            <Item>
                                <Label>
                                <Supertip>
                                    <Title>
                                    <Description>
                                <Action>
                                    <TaskpaneId>
                                    <SourceLocation>
                                    <Title>
                                    <FunctionName>
                                    <SourceLocation>
                <SourceLocation>
                <Label>
                <CommandSurface>
    <Resources>
        <Images>
            <Image>
                <Override>
        <Urls>
            <Url>
                <Override>
        <ShortStrings>
            <String>
                <Override>
        <LongStrings>
            <String>
                <Override>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="content-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="65d6f-131">VersionOverrides 内でのコンテンツ アドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="65d6f-131">Content add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a><span data-ttu-id="65d6f-132">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="65d6f-132">See also</span></span>

- [<span data-ttu-id="65d6f-133">アドイン Officeのリファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="65d6f-133">Reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="65d6f-134">公式スキーマの定義</span><span class="sxs-lookup"><span data-stu-id="65d6f-134">Official schema definitions</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
