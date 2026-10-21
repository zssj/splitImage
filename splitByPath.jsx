/*
splitByPath.jsx
用法：
1. 在图片上用钢笔/形状工具画若干闭合路径（每条路径 = 一块要切出的区域）
2. 同时选中这些路径 + 那张图片
3. 运行本脚本
结果：图片被按路径裁切为若干小块（原位保留），并各自栅格化为独立位图，
      原图和路径被删除。文件体积远小于“每块都带完整原图的剪切蒙版”方案。
支持嵌入图（RasterItem）和链接图（PlacedItem，运行时自动嵌入）。
*/

#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("请先打开文档。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert("请选中一条或多条路径，以及一张图片。");
        return;
    }

    var imageItem = null;
    var placedItem = null; // 链接图，稍后自动嵌入
    var pathItems = [];

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename === "RasterItem") {
            if (imageItem) {
                alert("只能选中一张图片。");
                return;
            }
            imageItem = item;
        } else if (item.typename === "PlacedItem") {
            // 链接图：先记录，收集完后统一 embed 成 RasterItem
            if (imageItem || placedItem) {
                alert("只能选中一张图片。");
                return;
            }
            placedItem = item;
        } else if (item.typename === "PathItem") {
            if (!item.closed) {
                item.closed = true;
            }
            pathItems.push(item);
        } else if (item.typename === "CompoundPathItem") {
            pathItems.push(item);
        }
    }

    // 链接图自动嵌入（方案A）
    if (!imageItem && placedItem) {
        // embed 前记：父容器 + 边界（embed 后原 PlacedItem 引用可能失效）
        var embedParent = null;
        var embedBounds = null;
        try { embedParent = placedItem.parent; } catch (e0) {}
        try { embedBounds = placedItem.geometricBounds; } catch (e1) {}

        try {
            placedItem.embed();
            try { app.redraw(); } catch (eDr) {}
        } catch (eEmbed) {
            alert("链接图嵌入失败，可能文件丢失或不可读：\n" + String(eEmbed));
            return;
        }

        // embed 后尝试沿用原引用
        var candidate = null;
        try {
            if (placedItem.typename === "RasterItem") candidate = placedItem;
        } catch (e2) {}

        // 引用失效时：在原父容器里按边界找新的 RasterItem
        if (!candidate && embedParent && embedBounds) {
            try {
                var pl = embedParent.pageItems;
                for (var zi = 0; zi < pl.length; zi++) {
                    var cand = pl[zi];
                    if (cand.typename !== "RasterItem") continue;
                    if (boundsAlmostEqual(cand.geometricBounds, embedBounds)) {
                        candidate = cand;
                        break;
                    }
                }
            } catch (e3) {}
        }

        if (!candidate) {
            alert("链接图已嵌入，但未能定位对应的图片对象，请手动选中嵌入图再运行。");
            return;
        }
        imageItem = candidate;
    }

    if (!imageItem) {
        alert("选区中没有找到图片（嵌入图或链接图均可）。");
        return;
    }

    if (pathItems.length === 0) {
        alert("选区中没有找到路径。");
        return;
    }

    // 计算原图有效 ppi（在删除原图之前，只算一次）
    var effectivePPI = getEffectivePPI(imageItem);
    // 记原图边界，删除失败时按边界补删
    var origImageBounds = null;
    try { origImageBounds = imageItem.geometricBounds; } catch (eOB) {}

    // 记录切割用路径，最后统一删除
    var pathsToDelete = pathItems.slice(0);

    // 栅格化选项（借鉴 preflight.jsx 已验证的 doc.rasterize 路径）
    var rasterOptions = new RasterizeOptions();
    rasterOptions.resolution = effectivePPI;
    rasterOptions.colorModel = RasterizationColorModel.DEFAULTCOLORMODEL;
    rasterOptions.transparency = true;          // 透明背景（保留路径外透明）
    rasterOptions.convertSpotColors = true;
    rasterOptions.includeLayers = false;
    rasterOptions.padding = 0.0;
    rasterOptions.artBoardClipping = false;     // 按对象自身边界，不按画板

    // 逐条路径切割 + 栅格化
    var failMsgs = [];
    for (var p = 0; p < pathItems.length; p++) {
        var cutPath = pathItems[p];

        // 复制图片和路径
        var imgCopy = imageItem.duplicate();
        var pathCopy = cutPath.duplicate();

        // 确保路径在图片上方（放到图层最前）
        pathCopy.move(doc.activeLayer, ElementPlacement.PLACEATBEGINNING);

        // 建立剪切蒙版：路径在上，图片在下，两者成组后设 clipping
        var group = doc.activeLayer.groupItems.add();
        pathCopy.move(group, ElementPlacement.PLACEATBEGINNING);
        imgCopy.move(group, ElementPlacement.PLACEATEND);
        group.clipped = true;

        // 栅格化范围必须用路径自身的 bounds
        // 不能用 group.geometricBounds —— 剪切组的 bounds 常含被裁掉的整图区域
        var clipBounds = null;
        try { clipBounds = cutPath.geometricBounds; } catch (eB) {}
        if (!clipBounds) {
            try { clipBounds = pathCopy.geometricBounds; } catch (eB2) {}
        }
        if (!clipBounds) {
            try { clipBounds = group.geometricBounds; } catch (eB3) {}
        }

        // 栅格化：doc.rasterize(item, bounds, options)
        // 必须显式传范围，传 null 会按画板范围栅格化（preflight v8.8.1 经验）
        try {
            var ni = doc.rasterize(group, clipBounds, rasterOptions);
            // 部分版本原对象不会自动移除
            try { group.remove(); } catch (eGroup) {}
        } catch (eRaster) {
            // 栅格化失败时保留剪切组（至少功能可用），并记录错误
            failMsgs.push("第 " + (p + 1) + " 块: " + String(eRaster));
        }
    }

    // 删除原图和原路径
    for (var d = 0; d < pathsToDelete.length; d++) {
        try {
            pathsToDelete[d].remove();
        } catch (eDel) {}
    }
    // 删除原图（引用失效时按边界补删）
    var imgRemoved = false;
    try {
        imageItem.remove();
        imgRemoved = true;
    } catch (eImg) {
        imgRemoved = false;
    }
    if (!imgRemoved && origImageBounds) {
        try {
            // 仅在当前图层顶层找与原图边界重合的 RasterItem（不递归，避免误删切块）
            var plTop = doc.activeLayer.pageItems;
            for (var bi = plTop.length - 1; bi >= 0; bi--) {
                var bit = plTop[bi];
                if (bit.typename !== "RasterItem") continue;
                if (boundsAlmostEqual(bit.geometricBounds, origImageBounds)) {
                    bit.remove();
                    break;
                }
            }
        } catch (eIB) {}
    }

    var msg = "分割完成：共切出 " + pathItems.length + " 块。\n分辨率：" + effectivePPI + " ppi（按原图有效分辨率）。";
    if (failMsgs.length > 0) {
        msg += "\n\n以下块栅格化失败（保留为剪切蒙版）：\n" + failMsgs.join("\n");
    }
    alert(msg);
})();

// 计算 RasterItem 有效 ppi（借鉴 preflight.jsx 的 pfPpiOf，用矩阵缩放反推）
// 不依赖 imageWidth/imageLength（部分版本取不到）
function getEffectivePPI(item) {
    var ppi = 0;
    try {
        var m = item.matrix;
        // 缩放因子取矩阵基向量模长：x轴→(A,B)，y轴→(C,D)，与旋转/斜切无关
        var sx = Math.sqrt(m.mValueA * m.mValueA + m.mValueB * m.mValueB);
        var sy = Math.sqrt(m.mValueC * m.mValueC + m.mValueD * m.mValueD);
        if (sx > 0.0001 && sy > 0.0001) {
            ppi = Math.min(72 / sx, 72 / sy); // 取较小值，保守
        }
    } catch (e) {}

    // 兜底：算不出来就用文档预设，绝不允许 NaN
    if (!ppi || isNaN(ppi) || ppi <= 0) {
        try { ppi = Number(app.activeDocument.rasterEffectSettings.resolution); } catch (e5) {}
    }
    if (!ppi || isNaN(ppi) || ppi <= 0) ppi = 300;

    // 限定在常用范围，避免极端值
    if (ppi < 72) ppi = 72;
    if (ppi > 2400) ppi = 2400;
    return Math.round(ppi);
}

// 比较两个 bounds 数组 [y1,x1,y2,x2] 是否几乎相等
function boundsAlmostEqual(a, b) {
    if (!a || !b) return false;
    var tol = 0.5;
    for (var i = 0; i < 4; i++) {
        if (Math.abs(Number(a[i]) - Number(b[i])) > tol) return false;
    }
    return true;
}
