---
layout: LandingPage
ms.topic: landing-page
title: Office JavaScript API リファレンス ドキュメント
description: Office JavaScript Api について説明します。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 8c467003bfd3481f35864225876e3ee024910eef
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234157"
---
# <a name="api-reference-documentation"></a>API リファレンス ドキュメント

アドインは Office JavaScript API を使用することで、Office クライアント アプリケーション内のオブジェクトを操作できます。 

<ul>
    <li><b>アプリケーション固有</b> API では、特定の Office アプリケーションにネイティブなオブジェクトを操作するために使用できる、厳密に型指定されたオブジェクトが提供されます。</li>
    <li><b>共通 API</b> を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</li>
</ul>

可能な場合は常にアプリケーション固有 API を使用するようにし、アプリケーション固有 API でサポートされていないシナリオに対してのみ共通 API を使用するようにします。 これらの 2 つの API モデルの詳細については、「<a href="../develop/develop-overview.md#api-models">Office アドインを開発する</a>」を参照してください。

<h2>API リファレンス</h2>

<ul class="panelContent cardsF cols cols3">
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/excel"><img src="../images/index/logo-excel.svg" alt="Excel API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Excel API リファレンス</h3>
                        <p><a href="/javascript/api/excel">Excel アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Outlook API リファレンス</h3>
                        <p><a href="/javascript/api/outlook">Outlook アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/word"><img src="../images/index/logo-word.svg" alt="Word API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Word API リファレンス</h3>
                        <p><a href="/javascript/api/word">Word アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>PowerPoint API リファレンス</h3>
                        <p><a href="/javascript/api/powerpoint">PowerPoint アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>OneNote API リファレンス</h3>
                        <p><a href="/javascript/api/onenote">OneNote アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/office"><img src="../images/index-landing-page/i_code-blocks.svg" alt="reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>共通 API リファレンス</h3>
                        <p><a href="/javascript/api/office">すべての Office アドインで使用できる JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>

<b>注</b>: 現在、Project 用のアプリケーション固有 JavaScript API はありません。Project アドインを構築する場合は、共通 API を使用してください。また、PowerPoint 用のアプリケーション固有 API の範囲は非常に限定的であるため、PowerPoint アドインを構築する際は、主に共通 API を使用してください。
