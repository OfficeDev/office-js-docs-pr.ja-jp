---
title: マニフェスト要素の正しい順序を確認する方法
description: 親要素内で子要素を配置するための正しい順序を確認する方法について説明します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 35ed1b87162b84ff13cafc2084ce9ca1b1666235
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087925"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="74c5c-103">マニフェスト要素の正しい順序を確認する方法</span><span class="sxs-lookup"><span data-stu-id="74c5c-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="74c5c-104">Office アドインのマニフェストの XML 要素は適切な親要素の下に配置する必要があり、*また*、親要素の下で子要素同士が特定の順序に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="74c5c-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="74c5c-105">必要な順序は、[[スキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)] フォルダー内の XSD ファイルで指定されています。</span><span class="sxs-lookup"><span data-stu-id="74c5c-105">The required ordering is specified in the XSD files in the [Schemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) folder.</span></span> <span data-ttu-id="74c5c-106">XSD ファイルは、作業ウィンドウ、コンテンツ、およびメール アドインのサブフォルダーに分類されます。</span><span class="sxs-lookup"><span data-stu-id="74c5c-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="74c5c-107">例えば、`<OfficeApp>` 要素では、`<Id>`、`<Version>`、`<ProviderName>` はこの順序で表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="74c5c-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="74c5c-108">`<AlternateId>` 要素が追加された場合、この要素は `<Id>` 要素と `<Version>` 要素の間に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="74c5c-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="74c5c-109">順序が間違っている要素が 1 つでもあると、マニフェストは有効にならず、アドインも読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="74c5c-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="74c5c-110">要素が間違った親の下にある場合とは異なり、 [office アドインマニフェスト内のバリデーター](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) は、要素の順序が間違っている場合に同じエラーメッセージを使用します。</span><span class="sxs-lookup"><span data-stu-id="74c5c-110">The [validator within office-addin-manifest](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="74c5c-111">エラーには、子要素が親要素の有効な子ではないと表示されます。</span><span class="sxs-lookup"><span data-stu-id="74c5c-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="74c5c-112">そのようなエラーが表示されるものの、子要素のレファレンス ドキュメントがこの子要素は親要素の有効な子 *である* と示す場合は、おそらく、子要素が間違った順序で配置されていることが原因です。</span><span class="sxs-lookup"><span data-stu-id="74c5c-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="74c5c-113">次のセクションでは、マニフェスト要素を表示する順序で示します。</span><span class="sxs-lookup"><span data-stu-id="74c5c-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="74c5c-114">`type`要素の属性が、、、のいずれであるかによって、相違点があり `<OfficeApp>` `TaskPaneApp` `ContentApp` `MailApp` ます。</span><span class="sxs-lookup"><span data-stu-id="74c5c-114">There are differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="74c5c-115">これらのセクションの扱いが大きくなりすぎないようにするため、非常に複雑な `<VersionOverrides>` 要素が別々のセクションに分割されます。</span><span class="sxs-lookup"><span data-stu-id="74c5c-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="74c5c-116">表示されている要素の一部は必須ではありません。</span><span class="sxs-lookup"><span data-stu-id="74c5c-116">Not all of the elements shown are mandatory.</span></span> <span data-ttu-id="74c5c-117">`minOccurs`[スキーマ](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3)で要素の値が **0** の場合、この要素は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="74c5c-117">If the `minOccurs` value for a element is **0** in the [schema](/openspecs/office_file_formats/ms-owemxml/4e112d0a-c8ab-46a6-8a6c-2a1c1d1299e3), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="74c5c-118">基本的な作業ウィンドウアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="74c5c-118">Basic task pane add-in element ordering</span></span>

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

<span data-ttu-id="74c5c-119">\*VersionOverrides の子要素の順序については、 [versionoverrides 内の作業ウィンドウアドイン要素の順序](#task-pane-add-in-element-ordering-within-versionoverrides) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="74c5c-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="74c5c-120">基本的なメールアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="74c5c-120">Basic mail add-in element ordering</span></span>

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

<span data-ttu-id="74c5c-121">\*Versionoverrides の子要素の順序については、「 [versionoverrides のメールアドイン要素の順序](#mail-add-in-element-ordering-within-versionoverrides-ver-10) 」と「1.0」および「 [メールアドイン1.1 要素](#mail-add-in-element-ordering-within-versionoverrides-ver-11) の順序」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="74c5c-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="74c5c-122">基本的なコンテンツアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="74c5c-122">Basic content add-in element ordering</span></span>

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

<span data-ttu-id="74c5c-123">\*VersionOverrides の子要素の順序については、 [versionoverrides 内のコンテンツアドイン要素の順序](#content-add-in-element-ordering-within-versionoverrides) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="74c5c-123">\*See [Content add-in element ordering within VersionOverrides](#content-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="74c5c-124">VersionOverrides 内の作業ウィンドウアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="74c5c-124">Task pane add-in element ordering within VersionOverrides</span></span>

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
                        <Group> (can be below <ControlGroup>)
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="74c5c-125">VersionOverrides 内のメールアドイン要素の順序は Ver です。</span><span class="sxs-lookup"><span data-stu-id="74c5c-125">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="74c5c-126">1.0</span><span class="sxs-lookup"><span data-stu-id="74c5c-126">1.0</span></span>

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

<span data-ttu-id="74c5c-127">\* の代わりに、値を指定した VersionOverrides は `type` `VersionOverridesV1_1` `VersionOverridesV1_0` 、外部 versionoverrides の末尾にネストすることができます。</span><span class="sxs-lookup"><span data-stu-id="74c5c-127">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="74c5c-128">の要素の順序については、「 [VersionOverrides overrides でのメールアドイン要素の順序](#mail-add-in-element-ordering-within-versionoverrides-ver-11)」を参照してください。 `VersionOverridesV1_1`</span><span class="sxs-lookup"><span data-stu-id="74c5c-128">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="74c5c-129">VersionOverrides 内のメールアドイン要素の順序は Ver です。</span><span class="sxs-lookup"><span data-stu-id="74c5c-129">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="74c5c-130">1.1</span><span class="sxs-lookup"><span data-stu-id="74c5c-130">1.1</span></span>

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

## <a name="content-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="74c5c-131">VersionOverrides 内でのコンテンツアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="74c5c-131">Content add-in element ordering within VersionOverrides</span></span>

```xml
<VersionOverrides>
    <WebApplicationInfo>
        <Id>
        <Resource>
        <Scopes>
            <Scope>
```

## <a name="see-also"></a><span data-ttu-id="74c5c-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="74c5c-132">See also</span></span>

- [<span data-ttu-id="74c5c-133">Office アドイン マニフェストのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="74c5c-133">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
