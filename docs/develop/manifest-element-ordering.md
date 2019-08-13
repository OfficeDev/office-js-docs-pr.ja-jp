---
title: マニフェスト要素の正しい順序を確認する方法
description: 親要素内で子要素を配置するための正しい順序を確認する方法について説明します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: d418f796592a0e4c247e717a5ce75d1c40c18d79
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302575"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="d8f8e-103">マニフェスト要素の正しい順序を確認する方法</span><span class="sxs-lookup"><span data-stu-id="d8f8e-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="d8f8e-104">Office アドインのマニフェストの XML 要素は適切な親要素の下に配置する必要があり、*また*、親要素の下で子要素同士が特定の順序に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="d8f8e-105">必要な順序は、[[スキーマ](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)] フォルダー内の XSD ファイルで指定されています。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-105">The required ordering is specified in the XSD files in the [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) folder.</span></span> <span data-ttu-id="d8f8e-106">XSD ファイルは、作業ウィンドウ、コンテンツ、およびメール アドインのサブフォルダーに分類されます。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="d8f8e-107">例えば、`<OfficeApp>` 要素では、`<Id>`、`<Version>`、`<ProviderName>` はこの順序で表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="d8f8e-108">`<AlternateId>` 要素が追加された場合、この要素は `<Id>` 要素と `<Version>` 要素の間に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="d8f8e-109">順序が間違っている要素が 1 つでもあると、マニフェストは有効にならず、アドインも読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="d8f8e-110">[Office-ツールボックス内のバリデーター](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-toolbox)は、要素が不適切な親の下にある場合と同じエラーメッセージを使用します。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-110">The [validator within office-toolbox](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-toolbox) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="d8f8e-111">エラーには、子要素が親要素の有効な子ではないと表示されます。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="d8f8e-112">そのようなエラーが表示されるものの、子要素のレファレンス ドキュメントがこの子要素は親要素の有効な子*である*と示す場合は、おそらく、子要素が間違った順序で配置されていることが原因です。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="d8f8e-113">次のセクションでは、マニフェスト要素を表示する順序で示します。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-113">The following sections show the manifest elements in the order in which they must appear.</span></span> <span data-ttu-id="d8f8e-114">`<OfficeApp>`要素の`type`属性が`TaskPaneApp`、 `ContentApp`、、のいずれであるかによって、 `MailApp`若干の違いがあります。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-114">There are slight differences depending on whether the `type` attribute of the `<OfficeApp>` element is `TaskPaneApp`, `ContentApp`, or `MailApp`.</span></span> <span data-ttu-id="d8f8e-115">これらのセクションの扱いが大きくなりすぎないように`<VersionOverrides>`するため、非常に複雑な要素が別々のセクションに分割されます。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-115">To keep these sections from becoming too unwieldy, the highly complex `<VersionOverrides>` element is broken out into separate sections.</span></span>

> [!Note]
> <span data-ttu-id="d8f8e-116">すべての要素が表示されるわけではありません。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-116">Not all of the elements show are mandatory.</span></span> <span data-ttu-id="d8f8e-117">スキーマで`minOccurs`要素の値が**0**の場合、 [](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)この要素は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-117">If the `minOccurs` value for a element is **0** in the [schema](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), the element is optional.</span></span>

## <a name="basic-task-pane-add-in-element-ordering"></a><span data-ttu-id="d8f8e-118">基本的な作業ウィンドウアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="d8f8e-118">Basic task pane add-in element ordering</span></span>

```
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
```

<span data-ttu-id="d8f8e-119">\*VersionOverrides の子要素の順序については、 [versionoverrides 内の作業ウィンドウアドイン要素の順序](#task-pane-add-in-element-ordering-within-versionoverrides)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-119">\*See [Task pane add-in element ordering within VersionOverrides](#task-pane-add-in-element-ordering-within-versionoverrides) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-mail-add-in-element-ordering"></a><span data-ttu-id="d8f8e-120">基本的なメールアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="d8f8e-120">Basic mail add-in element ordering</span></span>

```
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

<span data-ttu-id="d8f8e-121">\*Versionoverrides の子要素の順序については、「VersionOverrides の[メールアドイン要素の順序](#mail-add-in-element-ordering-within-versionoverrides-ver-10)」と「1.0」および「[メールアドイン1.1 要素](#mail-add-in-element-ordering-within-versionoverrides-ver-11)の順序」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-121">\*See [Mail add-in element ordering within VersionOverrides Ver. 1.0](#mail-add-in-element-ordering-within-versionoverrides-ver-10) and [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of children elements of VersionOverrides.</span></span>

## <a name="basic-content-add-in-element-ordering"></a><span data-ttu-id="d8f8e-122">基本的なコンテンツアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="d8f8e-122">Basic content add-in element ordering</span></span>

```
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
    <VersionOverrides>
```

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a><span data-ttu-id="d8f8e-123">VersionOverrides 内の作業ウィンドウアドイン要素の順序付け</span><span class="sxs-lookup"><span data-stu-id="d8f8e-123">Task pane add-in element ordering within VersionOverrides</span></span>

```
<VersionOverrides>
    <Description>
    <Requirements>
        <Sets>
            <Set>
      <Hosts>
        <Host>
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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a><span data-ttu-id="d8f8e-124">VersionOverrides 内のメールアドイン要素の順序は Ver です。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-124">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="d8f8e-125">1.0</span><span class="sxs-lookup"><span data-stu-id="d8f8e-125">1.0</span></span>

```
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

<span data-ttu-id="d8f8e-126">\*の`VersionOverridesV1_0`代わりに、 `type`値`VersionOverridesV1_1`を指定した VersionOverrides は、外部 versionoverrides の末尾にネストすることができます。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-126">\* A VersionOverrides with `type` value `VersionOverridesV1_1`, instead of `VersionOverridesV1_0`, can be nested at the end of the outer VersionOverrides.</span></span> <span data-ttu-id="d8f8e-127">の要素`VersionOverridesV1_1`の順序については、「 [versionoverrides Overrides でのメールアドイン要素の順序](#mail-add-in-element-ordering-within-versionoverrides-ver-11)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-127">See [Mail add-in element ordering within VersionOverrides Ver. 1.1](#mail-add-in-element-ordering-within-versionoverrides-ver-11) for the ordering of elements in `VersionOverridesV1_1`.</span></span>

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a><span data-ttu-id="d8f8e-128">VersionOverrides 内のメールアドイン要素の順序は Ver です。</span><span class="sxs-lookup"><span data-stu-id="d8f8e-128">Mail add-in element ordering within VersionOverrides Ver.</span></span> <span data-ttu-id="d8f8e-129">1.1</span><span class="sxs-lookup"><span data-stu-id="d8f8e-129">1.1</span></span>

```
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

## <a name="see-also"></a><span data-ttu-id="d8f8e-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="d8f8e-130">See also</span></span>

- [<span data-ttu-id="d8f8e-131">Office アドイン マニフェストのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="d8f8e-131">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
