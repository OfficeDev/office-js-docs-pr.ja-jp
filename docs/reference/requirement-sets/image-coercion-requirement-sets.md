---
title: 画像強制型変換要件セット
description: 複数のアドインを使用した Image Coercion 要件セットOffice、Excel、Word PowerPointサポート。
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 0f0b80c0af8213eaa9e3695373ddc037c2e60cc3
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/19/2021
ms.locfileid: "59448713"
---
# <a name="image-coercion-requirement-sets"></a>画像強制型変換要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="imagecoercion-11"></a>ImageCoercion 1.1

ImageCoercion 1.1 では、メソッドを使用してデータを書き込むときにイメージ ( `Office.CoercionType.Image` ) への変換が有効 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) です。 次のアプリケーションがサポートされています。

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

ImageCoercion 1.2 では、メソッドを使用してデータを書き込むときに SVG 形式 ( `Office.CoercionType.XmlSvg` ) に変換 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) できます。 次のアプリケーションがサポートされています。

- Excel 2021 以降のWindows
- Excel 2021 以降
- PowerPoint 2021 以降のWindows
- PowerPoint 2021 以降の Mac
- PowerPoint on the web
- Word 2021 以降のWindows
- Mac 上の Word 2021 以降

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
