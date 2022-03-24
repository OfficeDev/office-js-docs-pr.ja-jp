---
title: 画像強制型変換要件セット
description: 複数のアドインを使用した Image Coercion 要件セットOffice、Excel、Word PowerPointサポート。
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 50533d179180eeef81825a97c9c39fda95af554f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746921"
---
# <a name="image-coercion-requirement-sets"></a>画像強制型変換要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 を使用すると、メソッドを使用してデータを書き込むときにイメージ (`Office.CoercionType.Image`) への変換が可能 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) です。 次のアプリケーションがサポートされています。

- Excel 2013 以降のWindows
- Excel 2016以降の Mac
- Excel on iPad
- OneNote on the web
- PowerPoint 2013 以降のWindows
- PowerPoint 2016以降の Mac
- PowerPoint on the web
- PowerPoint on iPad
- Word on Windows (Word 2013 以降)
- Word on Mac (Word 2016 以降)
- Word on the web
- Word on iPad

## <a name="imagecoercion-12"></a>ImageCoercion 1.2

ImageCoercion 1.2 では、メソッドを使用してデータを書き込むときに SVG 形式 (`Office.CoercionType.XmlSvg`) に変換 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) できます。 次のアプリケーションがサポートされています。

- Excel 2021 以降のWindows
- Excel 2021 以降
- PowerPoint 2021 以降のWindows
- PowerPoint 2021 以降
- PowerPoint on the web
- Word 2021 以降のWindows
- Mac 上の Word 2021 以降

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
