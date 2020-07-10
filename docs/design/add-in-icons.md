---
title: Office アドインのアイコン ガイドライン
description: アドインコマンドのためのアイコンの設計方法と、最新のデザインスタイルおよび Monoline デザインスタイルの概要を説明します。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 35d8e0337b412a9ddebcde5be4db4db802e88269
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093841"
---
# <a name="icons"></a>アイコン

アイコンは、動作や概念を視覚的に表現するものです。 多くの場合、コントロールとコマンドに意味を与えるために使用します。 環境内でユーザーが移動するのにサインが役立つのと同じように、リアルなビジュアルや象徴的なビジュアルにより、ユーザーは UI 間を移動できるようになります。 お客様がコントロールを選択するときの動作をすばやく解析できるにように、必要な詳細のみを含む、シンプルで明確なビジュアルにする必要があります。

Office アプリのリボンインターフェイスには、標準の視覚スタイルがあります。 これにより、Office アプリとの間で一貫性と親和性を保つことができます。 このガイドラインは、ソリューションの PNG アセットのセットを Office の自然な一部のように設計するのに役立ちます。

Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.

## <a name="design-icons-for-add-in-commands"></a>アドイン コマンドのアイコンをデザインする

[アドイン コマンド](add-in-commands.md)は、Office UI にボタン、テキスト、およびアイコンを追加します。 アドイン コマンドのボタンには、ユーザーがコマンドを使うときに、実行しようとするアクションを明確に識別できる、分かりやすいアイコンとラベルをつける必要があります。 次の記事では、Office とシームレスに統合されるアイコンを設計するのに役立つスタイルと運用上のガイドラインを提供します。

- Microsoft 365 の Monoline スタイルについては、「 [Office アドインの Monoline スタイルアイコンガイドライン](add-in-icons-monoline.md)」を参照してください。
- サブスクリプション以外の Office 2013 以降の新しいスタイルについては、「 [Office アドインの新しいスタイルのアイコンガイドライン](add-in-icons-fresh.md)」を参照してください。

> [!NOTE]
> どちらか一方のスタイルを選択する必要があります。また、アドインでは、Microsoft 365 またはサブスクリプション以外の Office で実行されている場合と同じアイコンを使用します。

## <a name="see-also"></a>関連項目

- [アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [Excel、Word、PowerPoint のアドイン コマンド](../design/add-in-commands.md)
