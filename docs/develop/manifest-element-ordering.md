---
title: マニフェスト要素の正しい順序を確認する方法
description: 親要素内で子要素を配置するための正しい順序を確認する方法について説明します。
ms.date: 08/15/2019
localization_priority: Normal
ms.openlocfilehash: 68eaa21af9a86c910a9e8701ef2ea6cf95a64000
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477762"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>マニフェスト要素の正しい順序を確認する方法

Office アドインのマニフェストの XML 要素は適切な親要素の下に配置する必要があり、*また*、親要素の下で子要素同士が特定の順序に配置する必要があります。

必要な順序は、[[スキーマ](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)] フォルダー内の XSD ファイルで指定されています。 XSD ファイルは、作業ウィンドウ、コンテンツ、およびメール アドインのサブフォルダーに分類されます。

例えば、`<OfficeApp>` 要素では、`<Id>`、`<Version>`、`<ProviderName>` はこの順序で表示する必要があります。 `<AlternateId>` 要素が追加された場合、この要素は `<Id>` 要素と `<Version>` 要素の間に配置する必要があります。 順序が間違っている要素が 1 つでもあると、マニフェストは有効にならず、アドインも読み込まれません。

> [!NOTE]
> 要素が間違った親の下にある場合とは異なり、 [office アドインマニフェスト内のバリデーター](../testing/troubleshoot-manifest.md#validate-your-manifest-with-office-addin-manifest)は、要素の順序が間違っている場合に同じエラーメッセージを使用します。 エラーには、子要素が親要素の有効な子ではないと表示されます。 そのようなエラーが表示されるものの、子要素のレファレンス ドキュメントがこの子要素は親要素の有効な子*である*と示す場合は、おそらく、子要素が間違った順序で配置されていることが原因です。

次のセクションでは、マニフェスト要素を表示する順序で示します。 `<OfficeApp>`要素の`type`属性が`TaskPaneApp`、 `ContentApp`、、のいずれであるかによって、 `MailApp`若干の違いがあります。 これらのセクションの扱いが大きくなりすぎないように`<VersionOverrides>`するため、非常に複雑な要素が別々のセクションに分割されます。

> [!Note]
> すべての要素が表示されるわけではありません。 スキーマで`minOccurs`要素の値が**0**の場合、 [](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)この要素は省略可能です。

## <a name="basic-task-pane-add-in-element-ordering"></a>基本的な作業ウィンドウアドイン要素の順序付け

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

\*VersionOverrides の子要素の順序については、 [versionoverrides 内の作業ウィンドウアドイン要素の順序](#task-pane-add-in-element-ordering-within-versionoverrides)を参照してください。

## <a name="basic-mail-add-in-element-ordering"></a>基本的なメールアドイン要素の順序付け

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

\*Versionoverrides の子要素の順序については、「VersionOverrides の[メールアドイン要素の順序](#mail-add-in-element-ordering-within-versionoverrides-ver-10)」と「1.0」および「[メールアドイン1.1 要素](#mail-add-in-element-ordering-within-versionoverrides-ver-11)の順序」を参照してください。

## <a name="basic-content-add-in-element-ordering"></a>基本的なコンテンツアドイン要素の順序付け

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

## <a name="task-pane-add-in-element-ordering-within-versionoverrides"></a>VersionOverrides 内の作業ウィンドウアドイン要素の順序付け

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

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-10"></a>VersionOverrides 内のメールアドイン要素の順序は Ver です。 1.0

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

\*の`VersionOverridesV1_0`代わりに、 `type`値`VersionOverridesV1_1`を指定した VersionOverrides は、外部 versionoverrides の末尾にネストすることができます。 の要素`VersionOverridesV1_1`の順序については、「 [versionoverrides Overrides でのメールアドイン要素の順序](#mail-add-in-element-ordering-within-versionoverrides-ver-11)」を参照してください。

## <a name="mail-add-in-element-ordering-within-versionoverrides-ver-11"></a>VersionOverrides 内のメールアドイン要素の順序は Ver です。 1.1

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

## <a name="see-also"></a>関連項目

- [Office アドイン マニフェストのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
