---
title: マニフェスト ファイルの Permissions 要素
description: Permissions 要素は、アドインの API アクセス レベルOffice指定します。
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: 2f2ccb4f6ec691b19cadea76a06520a9bad7a0b6c0e51699f2c8db67a3030de0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089014"
---
# <a name="permissions-element"></a>Permissions 要素

Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

コンテンツ アドインおよび作業ウィンドウ アドインの場合:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

メール アドインの場合:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

詳細については、「コンテンツアドインと作業ウィンドウ アドインでの[API](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)使用のアクセス許可の要求」および「Outlookについて」を[参照してください](../../outlook/understanding-outlook-add-in-permissions.md)。
