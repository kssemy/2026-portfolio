// ============================================
// Export > Visible Layers to Files
// 눈 켜진 레이어만 레이어 이름 그대로 JPG 저장
// 포토샵 기본 기능의 번호 접두사 없이 레이어명으로만 저장
// 사용법: 파일 > 스크립트 > 찾아보기 로 실행
// ============================================

#target photoshop

var doc = app.activeDocument;
var originalState = doc.activeHistoryState;
var originalVisible = [];

// 저장 폴더 선택
var outputFolder = Folder.selectDialog("레이어를 저장할 폴더를 선택하세요");

if (outputFolder == null) {
    alert("폴더가 선택되지 않았습니다.");
} else {

    var JPG_QUALITY = 12;

    function sanitizeFileName(name) {
        return name
            .replace(/[\\\/\:\*\?\"\<\>\|]/g, "_")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "");
    }

    function collectLayers(layers, result) {
        for (var i = 0; i < layers.length; i++) {
            if (layers[i].typename === "LayerSet") {
                collectLayers(layers[i].layers, result);
            } else {
                result.push(layers[i]);
            }
        }
        return result;
    }

    function saveVisibility(layers, store) {
        for (var i = 0; i < layers.length; i++) {
            store.push(layers[i].visible);
            if (layers[i].typename === "LayerSet") {
                saveVisibility(layers[i].layers, store);
            }
        }
    }

    function hideAll(layers) {
        for (var i = 0; i < layers.length; i++) {
            layers[i].visible = false;
            if (layers[i].typename === "LayerSet") {
                hideAll(layers[i].layers);
            }
        }
    }

    function restoreVisibility(layers, store, idx) {
        for (var i = 0; i < layers.length; i++) {
            layers[i].visible = store[idx[0]];
            idx[0]++;
            if (layers[i].typename === "LayerSet") {
                restoreVisibility(layers[i].layers, store, idx);
            }
        }
    }

    var allLayers = collectLayers(doc.layers, []);
    saveVisibility(doc.layers, originalVisible);

    var exportLayers = [];
    for (var i = 0; i < allLayers.length; i++) {
        if (allLayers[i].visible) {
            exportLayers.push(allLayers[i]);
        }
    }

    var count = 0;
    var skipped = 0;

    for (var i = 0; i < exportLayers.length; i++) {
        var layer = exportLayers[i];

        try {
            hideAll(doc.layers);

            layer.visible = true;

            var parent = layer.parent;
            while (parent && parent.typename !== "Document") {
                parent.visible = true;
                parent = parent.parent;
            }

            var cleanName = sanitizeFileName(layer.name);
            if (cleanName === "") cleanName = "layer_" + (i + 1);

            var saveFile = new File(outputFolder + "/" + cleanName + ".jpg");
            var dupCount = 1;

            while (saveFile.exists) {
                saveFile = new File(outputFolder + "/" + cleanName + "_" + dupCount + ".jpg");
                dupCount++;
            }

            var jpgOptions = new JPEGSaveOptions();
            jpgOptions.quality = JPG_QUALITY;
            jpgOptions.embedColorProfile = true;
            jpgOptions.matte = MatteType.NONE;
            jpgOptions.formatOptions = FormatOptions.STANDARDBASELINE;

            doc.saveAs(saveFile, jpgOptions, true, Extension.LOWERCASE);
            count++;

        } catch (e) {
            skipped++;
        }
    }

    doc.activeHistoryState = originalState;
    restoreVisibility(doc.layers, originalVisible, [0]);

    alert(
        "✅ 저장 완료!\n\n" +
        "저장된 파일: " + count + "개\n" +
        "건너뜀: " + skipped + "개\n" +
        "처음에 눈 꺼져있던 레이어: " + (allLayers.length - exportLayers.length) + "개\n\n" +
        "📁 저장 위치:\n" + outputFolder.fsName
    );
}
