// 대용량 CAD 업로드 + NDJSON 스트림 수신 공용 헬퍼.
//
// remote30_prototype.html 안에 인라인으로 있던 것을 필드명·엔드포인트에 매이지 않게
// 일반화했다. 프로토타입은 아직 자기 사본을 쓴다 — 회귀 위험이 커서 함께 옮기지 않았다.
(function (global) {
  "use strict";

  // DXF 는 ASCII 라 gzip 이 8배 가까이 줄인다(실측: 110.6MB → 14.4MB). 서버
  // `_save_upload` 가 ".gz" 확장자와 매직바이트를 보고 알아서 해제한다.
  // CompressionStream 미지원 브라우저는 null → 호출측이 원본을 그대로 보낸다.
  async function gzipBlob(file, onProgress) {
    if (typeof CompressionStream === "undefined") return null;
    try {
      let source = file.stream();
      if (onProgress) {
        const total = file.size || 1;
        let read = 0;
        source = source.pipeThrough(new TransformStream({
          transform(chunk, ctrl) {
            read += chunk.byteLength;
            onProgress(Math.min(1, read / total));
            ctrl.enqueue(chunk);
          },
        }));
      }
      const stream = source.pipeThrough(new CompressionStream("gzip"));
      const buf = await new Response(stream).arrayBuffer();
      return new Blob([buf], { type: "application/gzip" });
    } catch (e) { return null; }
  }

  // 압축이 되레 커지는 파일(이미 압축된 DWG 등)은 원본을 보낸다.
  async function appendUpload(fd, field, file, onProgress) {
    const gz = await gzipBlob(file, onProgress);
    if (gz && gz.size < file.size) {
      fd.append(field, gz, file.name + ".gz");
      return gz.size;
    }
    fd.append(field, file, file.name);
    return file.size;
  }

  // 업로드 진행률은 XHR 로만 얻을 수 있다. 작은 JSON 응답 하나만 받는 업로드 전용
  // 요청으로 분리해, 진행률은 유지하면서 큰 응답을 XHR 에 태우지 않는다.
  function xhrUploadForToken(url, body, onUploadProgress) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", url, true);
      if (onUploadProgress) {
        xhr.upload.onprogress = (e) => {
          if (e.lengthComputable) onUploadProgress(e.loaded / e.total);
        };
      }
      xhr.onerror = () => reject(new Error("네트워크 오류로 업로드가 실패했습니다"));
      xhr.onabort = () => reject(new Error("업로드가 중단되었습니다"));
      xhr.onload = () => {
        let data = null;
        try { data = JSON.parse(xhr.responseText); } catch (e) {}
        if (xhr.status < 200 || xhr.status >= 300 || !data || !data.ok) {
          reject(new Error((data && data.message) || `HTTP ${xhr.status}`));
          return;
        }
        resolve(data);
      };
      xhr.send(body);
    });
  }

  async function readNdjson(res, onMessage) {
    const reader = res.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buf = "";
    const drain = (final) => {
      // 줄 단위로 한 번에 쪼갠다 — 줄마다 buf.slice 로 잘라내면 한 청크에 수백 줄이
      // 실릴 때 남은 버퍼를 매번 복사해 O(줄²) 이 된다.
      const lines = buf.split("\n");
      buf = lines.pop();
      if (final && buf.trim()) lines.push(buf);
      if (final) buf = "";
      for (let i = 0; i < lines.length; i++) {
        if (!lines[i]) continue;
        let msg;
        try { msg = JSON.parse(lines[i]); } catch (e) { continue; }
        onMessage(msg);
      }
    };
    try {
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        drain(false);
      }
      drain(true);
    } catch (e) {
      try { await reader.cancel(); } catch (_) {}
      throw e;
    }
  }

  async function readError(res) {
    let m = `HTTP ${res.status}`;
    try { m = (await res.json()).message || m; } catch (e) {}
    return new Error(m);
  }

  // 압축 → 업로드(토큰) → 파싱 스트림. 파일과 스트림을 한 요청에 묶지 않는 이유는
  // XHR 이 응답 전체를 responseText 에 붙들고 있어서다 — LH 지하층배관도의 inspect
  // NDJSON 은 비압축 173MB 라 그대로 탭 메모리에 남는다. 토큰만 XHR 로 받고 본
  // 스트림은 fetch + ReadableStream 으로 읽어 소비한 청크를 즉시 버린다.
  async function uploadAndStream(opts) {
    const { file, field, uploadUrl, streamUrl, extra, onPhase, onMessage } = opts;
    const phase = onPhase || function () {};

    const fd = new FormData();
    phase("compress", 0, { bytes: file.size });
    const sentBytes = await appendUpload(fd, field, file,
      (p) => phase("compress", p, { bytes: file.size }));

    phase("upload", 0, { bytes: sentBytes, original: file.size });
    const up = await xhrUploadForToken(uploadUrl, fd,
      (p) => phase("upload", p, { bytes: sentBytes, original: file.size }));

    const body = new FormData();
    if (extra) for (const [k, v] of Object.entries(extra)) body.append(k, v);
    body.set("dxf_token", up.dxf_token);

    phase("parse", 0, {});
    const res = await fetch(streamUrl, { method: "POST", body });
    if (!res.ok) throw await readError(res);
    await readNdjson(res, onMessage);
    return up;
  }

  global.UploadStream = {
    gzipBlob, appendUpload, xhrUploadForToken, readNdjson, uploadAndStream,
  };
})(window);
