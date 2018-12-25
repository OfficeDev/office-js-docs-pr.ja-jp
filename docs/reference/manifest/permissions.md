---
title: マニフェスト ファイルの Permissions 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 9193651ec0c795cdb55eb3fc6576dbacd59e0fb2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432355"
---
# <a name="permissions-element"></a>Permissions 要素

Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

コンテンツ アドインおよび作業ウィンドウ アドインの場合:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

メール アドインの場合

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>解説

詳細については、「[コンテンツ アドインおよび作業ウィンドウ アドインでの API 使用のアクセス許可を要求する](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)」と「[Outlook アドインのアクセス許可について](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」をご覧ください。
