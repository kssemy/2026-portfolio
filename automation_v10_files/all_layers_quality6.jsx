// ============================================
// Export > Layers to Files (레이어 이름 그대로 저장)
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

    // JPG 품질 설정 (1~12, 12가 최고 품질)
    var JPG_QUALITY = 6;

    // 파일명에 쓸 수 없는 특수문자 치환
    function sanitizeFileName(name) {
        return name.replace(/[\\\/\:\*\?\"\<\>\|]/g, "_").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    // 모든 레이어 (LayerSet 제외, 일반 레이어만) 수집
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

    // 현재 visible 상태 저장
    function saveVisibility(layers, store) {
        for (var i = 0; i < layers.length; i++) {
            store.push(layers[i].visible);
            if (layers[i].typename === "LayerSet") {
                saveVisibility(layers[i].layers, store);
            }
        }
    }

    // 전체 레이어 숨기기
    function hideAll(layers) {
        for (var i = 0; i < layers.length; i++) {
            layers[i].visible = false;
            if (layers[i].typename === "LayerSet") {
                hideAll(layers[i].layers);
            }
        }
    }

    // visible 상태 복원
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

    var count = 0;
    var skipped = 0;

    for (var i = 0; i < allLayers.length; i++) {
        var layer = allLayers[i];

        try {
            // 전체 숨기고 해당 레이어만 표시
            hideAll(doc.layers);
            layer.visible = true;

            // 부모 그룹이 있으면 표시
            var parent = layer.parent;
            while (parent && parent.typename !== "Document") {
                parent.visible = true;
                parent = parent.parent;
            }

            // 파일명 정리
            var cleanName = sanitizeFileName(layer.name);
            if (cleanName === "") cleanName = "layer_" + (i + 1);

            // 중복 파일명 처리
            var saveFile = new File(outputFolder + "/" + cleanName + ".jpg");
            var dupCount = 1;
            while (saveFile.exists) {
                saveFile = new File(outputFolder + "/" + cleanName + "_" + dupCount + ".jpg");
                dupCount++;
            }

            // JPG 저장 옵션
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

    // 원래 상태 복원
    doc.activeHistoryState = originalState;
    restoreVisibility(doc.layers, originalVisible, [0]);

    alert("✅ 저장 완료!\n\n저장된 파일: " + count + "개\n건너뜀: " + skipped + "개\n\n📁 저장 위치:\n" + outputFolder.fsName);
}
