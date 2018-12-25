---
title: マニフェスト ファイルの OfficeTab 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 721064687c3c892b565a94e418815726cc0817f5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432873"
---
# <a name="officetab-element"></a>OfficeTab 要素

アドイン コマンドを表示するリボン タブを定義します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。この要素は必須です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  Description  |
|:-----|:-----|:-----|
|  Group      | はい |  コマンドのグループを定義します。既定のタブには、アドインごとに 1 つのグループのみを追加できます。  |

ホストごとの有効なタブ `id` 値は次のとおりです。 **太字**の値は、デスクトップとオンラインの両方でサポートされています (たとえば、Windows 用の Word 2016 以降と Word Online)。

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

タブの UI 拡張ポイントのグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性は必須であり、各 **id** 属性はマニフェスト内で一意でなければなりません。**id** は最大 125 文字の文字列です。「[Group 要素](group.md)」を参照してください。

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
