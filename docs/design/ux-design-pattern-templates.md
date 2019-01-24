---
title: Office アドイン用 UX 設計パターン
description: ''
ms.date: 06/27/2018
localization_priority: Priority
ms.openlocfilehash: 65875bced876bb20243e3f2f0a353aa10c158a79
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387822"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Office アドイン用 UX 設計パターン

Office アドインのユーザー エクスペリエンスの設計では、Office ユーザーにとって魅力的なエクスペリエンスを提供し、既定の Office UI 内でシームレスに適合させることにより Office 全体のエクスペリエンスを拡張する必要があります。  

この UX パターンはコンポーネントで構成されています。 コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。 ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。

Office UI Fabric では、外観も動作も Office の一部のようなコンポーネントを表示します。 Fabric を活用して、Office と簡単に統合します。 アドインに既存のコンポーネント言語がある場合、Fabric のためにその言語を削除する必要はありません。 Office と統合する際に、それを保持する機会を探します。 スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルと動作の採用を行う方法を検討してください。

提供されるパターンは、一般的な顧客シナリオとユーザー エクスペリエンス調査に基づくベスト プラクティス ソリューションです。 それは、アドインの設計と開発のためのクイック エントリ ポイントと、Microsoft とブランド要素のバランスを実現するためのガイダンスの両方を提供することを目的としています。 Microsoft の Fabric 設計言語とパートナー特有のブランドの独自性から得たデザイン要素のバランスを取る、クリーンでモダンなユーザー エクスペリエンスを提供することにより、ユーザーの保持とアドインの採用を向上させることができます。

UX パターン テンプレートを使用して、次のことを行います。

* よくある顧客のシナリオにソリューションとして適用する。
* 設計のベスト プラクティスとして適用する。
* [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) のコンポーネントとスタイルを組み込む。
* Office の既定の UI に視覚的に溶け込むアドインをビルドする。
* UX を観念化および可視化する。


## <a name="getting-started"></a>はじめに

パターンは、キーの動作またはアドインに共通のエクスペリエンスによって構成されます。 主なグループは次のとおりです。

* [最初の実行エクスペリエンス (FRE)](../design/first-run-experience-patterns.md)
* [認証](../design/authentication-patterns.md)
* [ナビゲーション](../design/navigation-patterns.md)
* [ブランド デザイン](../design/branding-patterns.md)

各グループを参照して、ベスト プラクティスを使ってアドインを設計する方法を理解します。



>注: このドキュメント全体で示す画面の例は、解像度 **1366x768** で設計および表示されています。




## <a name="see-also"></a>関連項目
* [デザインのツールキット](design-toolkits.md)
* [Office UI Fabric](https://developer.microsoft.com/fabric)
* [Office アドイン開発のベスト プラクティス](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-best-practices)
* [Fabric React の使用の開始](https://docs.microsoft.com/office/dev/add-ins/design/using-office-ui-fabric-react)
