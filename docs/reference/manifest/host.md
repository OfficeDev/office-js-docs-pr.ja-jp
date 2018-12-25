---
title: マニフェスト ファイルの Host 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 37b772261ad82b4f899e73314a08ffd1dd03b442
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432810"
---
# <a name="host-element"></a>Host 要素

アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。

> [!IMPORTANT] 
> **Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。 ただし、機能は変わりません。  

## <a name="basic-manifest"></a>基本のマニフェスト

基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。   

### <a name="attributes"></a>属性

| 属性     | 種類   | 必須 | 説明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | 必須 | Office ホスト アプリケーションの種類の名前。 |

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
|  [MobileFormFactor](mobileformfactor.md)    |  いいえ   |  モバイル フォーム ファクターの設定を定義します。**注:** この要素は、Outlook for iOS でのみサポートされています。 |
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
