---
title: マニフェスト ファイルの Host 要素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127570"
---
# <a name="host-element"></a>Host 要素

アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。

> [!IMPORTANT] 
> **Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。 ただし、機能は変わりません。  

## <a name="basic-manifest"></a>基本のマニフェスト

基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。

### <a name="attributes"></a>属性

| 属性     | 型   | 必須 | 説明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [名前](#name) | string | 必須 | Office ホスト アプリケーションの種類の名前。 |

### <a name="name"></a>名前
このアドインが対象にするホストの種類を指定します。この値は、次のいずれかである必要があります。

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>例
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>VersionOverrides ノード
[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。 

### <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  はい  | これらの設定を適用する Office ホストについて説明します。|

### <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  はい   |  デスクトップ フォーム ファクターの設定を定義します。 |
|  [MobileFormFactor](mobileformfactor.md)    |  いいえ   |  モバイルフォームファクターの設定を定義します。 **注:** この要素は、iOS の Outlook でのみサポートされています。 |
|  [AllFormFactors](allformfactors.md)    |  いいえ   |  すべてのフォーム ファクターの設定を定義します。 Excel のカスタム関数でのみ使用します。 |

### <a name="xsitype"></a>xsi:type

含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。 この値は、次のいずれかである必要があります。

- `Document` (Word)
- `MailHost` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>ホストの例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
