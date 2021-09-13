---
title: マニフェスト ファイルの OfficeTab 要素
description: OfficeTab 要素は、アドイン コマンドが表示されるリボン タブを定義します。
ms.date: 06/20/2019
ms.localizationpriority: medium
ms.openlocfilehash: 94af0e8744c496538a506fdc4626cd3f515129bb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152969"
---
# <a name="officetab-element"></a>OfficeTab 要素

アドイン コマンドを表示するリボン タブを定義します。 これは、既定のタブ ([**ホーム**]、[**メッセージ**]、または [会議] のいずれか)、またはアドインによって定義されたカスタム タブのいずれかです。  この要素は必須です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  グループ      | はい |  コマンドのグループを定義します。 既定のタブには、アドインごとに 1 つのグループのみを追加できます。  |

アプリケーション別の有効なタブ `id` 値を次に示します。 太字の **値** は、デスクトップとオンラインの両方でサポートされます (たとえば、Word 2016以降のWindowsおよびWord on the web)。

### <a name="outlook"></a>Outlook

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>Group

タブ内の UI 拡張ポイントのグループ。グループには最大 6 つのコントロールを含めできます。 **id 属性は** 必須であり、各 **ID** はマニフェスト内で一意である必要があります。 id **は** 、最大 125 文字の文字列です。 [Group 要素](group.md)を参照してください。

## <a name="officetab-example"></a>OfficeTab の例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
