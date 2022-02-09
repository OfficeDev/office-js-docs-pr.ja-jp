---
title: マニフェスト ファイル内の MobileButton 型のコントロール要素
description: アクションを実行するか、作業ウィンドウを起動するモバイル デバイス上のボタンを定義します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d498b728bf7f19cf239ffc6178f19cdf9a62de58
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467933"
---
# <a name="control-element-of-type-mobilebutton"></a>MobileButton 型のコントロール要素

アクションを実行するか、作業ウィンドウを起動し、モバイル プラットフォームにのみ表示されるボタンを定義します。

> [!NOTE]
> この記事では、要素の属性に関する重要な情報を含む基本的な [Control](control.md) リファレンス記事に精通している必要があります。

モバイル ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。 関数を実行するか、作業ウィンドウを表示します。 各ボタン コントロールには、マニフェスト内 `id` のすべての Control 要素間で一意の **属性値** が必要です。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

`MobileButton` の  値は、VersionOverrides スキーマ 1.1 で定義されます。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)     | はい |  ボタンのテキストです。 |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。 Control 要素の Action 子 **は** 1 つ **のみです** 。 |

### <a name="label"></a>ラベル

ボタンのテキストを、32 文字以内にできる唯一の属性を使用して指定し、Resources 要素の **ShortStrings** 子の **String** 要素の **id** 属性の値に設定する [必要](resources.md)があります。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="examples"></a>例

次の例では、ボタンは関数を実行します。

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

次の例では、ボタンに作業ウィンドウが表示されます。

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
