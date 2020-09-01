---
title: Office アドインを開発する
description: Office アドイン開発の概要を説明します。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 419880e8872df20be5a3de40f480f70be2b18859
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292779"
---
# <a name="develop-office-add-ins"></a>Office アドインを開発する

> [!TIP]
> この記事を読む前に、「[Building Office Add-ins (Office アドインの構築)](../overview/office-add-ins-fundamentals.md)」をご覧ください。

すべての Office アドインは、Office アドイン プラットフォーム上で構築します。 すべての Office アドインでは共通のフレームワークが共有され、これにより特定の機能の実装が可能になります。 どのようなアドインを構築する場合でも、アプリケーションやプラットフォームの可用性、Office JavaScript API のプログラミング パターン、アドインの設定と機能をマニフェスト ファイル上で指定する方法など、重要な概念を理解する必要があります。 開発に関するこれらの中心概念については、ドキュメントの「**Core concepts (中心概念)**」 > 「**Develop (開発)**」セクションを参照してください。 構築するアドインに対応するアプリケーション固有のドキュメント (たとえば、 [Excel](../excel/index.yml)) を詳しく見る前に、ここに記載される情報を確認してください。

> [!NOTE]
> 「**Core concepts (中心概念)**」 > 「**Develop (開発)**」 > 「**How to (方法)**」セクションには、開発に関する具体的な概念やタスクについての記事があります。 同セクションでは、[Visual Studio Code を使用したアドイン開発](develop-add-ins-vscode.md)、[タスク ウィンドウをドキュメントと共に自動的に開く](automatically-open-a-task-pane-with-a-document.md)、[アドイン コマンドの作成](create-addin-commands.md)、[ダイアログ ボックスを開く](dialog-api-in-office-add-ins.md)などに関する情報が提供されています。

## <a name="next-steps"></a>次の手順

ここで説明する中心概念について理解したら、構築するアドインに対応するアプリケーション固有のドキュメント (たとえば、[Excel](../excel/index.yml)) を確認します。 ドキュメントの各アプリケーション固有のセクションには、特定の Office アプリケーション用のアドインの構築に関する具体的な情報が記載されています。

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Office アドインの構築](../overview/office-add-ins-fundamentals.md)
- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインの設計](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインを発行する](../publish/publish.md)
