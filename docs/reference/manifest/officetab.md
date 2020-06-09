---
title: マニフェスト ファイルの OfficeTab 要素
description: OfficeTab 要素は、アドインコマンドが表示されるリボンタブを定義します。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b4bfd83c210a1b0a8a443e1a3f849974124a6b61
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611513"
---
# <a name="officetab-element"></a>OfficeTab 要素

アドイン コマンドを表示するリボン タブを定義します。 これは、既定のタブ ([**ホーム**]、[**メッセージ**]、または [**会議**]) にするか、アドインで定義されたカスタムタブにすることができます。 この要素は必須です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  グループ      | はい |  コマンドのグループを定義します。 既定のタブには、アドインごとに 1 つのグループのみを追加できます。  |

ホストごとの有効なタブ `id` 値は次のとおりです。 **太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、word 2016 以降の Windows および web 上の word)。

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

タブ内の UI 拡張ポイントのグループ。グループは最大6つのコントロールを持つことができます。 **Id**属性は必須で、各**id**はマニフェスト内で一意である必要があります。 **Id**は、最大125文字の文字列です。 [Group 要素](group.md)を参照してください。

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
