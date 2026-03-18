/**
 * 웹사이트 스크린샷 캡처 / 이미지 다운로드 스크립트
 *
 * URL이 이미지 파일이면 직접 다운로드하고,
 * 웹사이트이면 Playwright를 사용하여 스크린샷을 캡처한다.
 *
 * Usage:
 *   node capture.mjs <URL> <저장경로> [옵션]
 *
 * Options:
 *   --full-page        전체 페이지 스크롤 캡처
 *   --width <number>   뷰포트 너비 (기본: 1920)
 *   --height <number>  뷰포트 높이 (기본: 1080)
 *   --wait <number>    로드 후 대기 시간(초) (기본: 2)
 *   --device <name>    모바일 디바이스 에뮬레이션
 */

import { chromium, devices } from "playwright";
import { mkdirSync, writeFileSync } from "fs";
import { dirname, resolve, extname } from "path";
import { parseArgs } from "util";

const IMAGE_EXTENSIONS = new Set([
  ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp", ".ico", ".tiff", ".avif",
]);

/**
 * URL이 이미지인지 판별한다.
 * 1) 확장자 기반 판별
 * 2) HTTP HEAD 요청의 Content-Type 판별
 */
async function isImageUrl(url) {
  // 확장자 판별 (쿼리스트링 제거 후)
  const pathname = new URL(url).pathname;
  const ext = extname(pathname).toLowerCase();
  if (IMAGE_EXTENSIONS.has(ext)) return true;

  // Content-Type 판별
  try {
    const res = await fetch(url, { method: "HEAD", redirect: "follow", signal: AbortSignal.timeout(10000) });
    const contentType = res.headers.get("content-type") || "";
    return contentType.startsWith("image/");
  } catch (e) {
    console.error(`HEAD check failed for ${url}: ${e.message}`);
    return false;
  }
}

/**
 * 이미지 URL을 직접 다운로드한다.
 */
async function downloadImage(url, outputPath) {
  const res = await fetch(url, { redirect: "follow", signal: AbortSignal.timeout(30000) });
  if (!res.ok) {
    throw new Error(`이미지 다운로드 실패: HTTP ${res.status} ${res.statusText}`);
  }
  const buffer = Buffer.from(await res.arrayBuffer());
  writeFileSync(outputPath, buffer);
  console.log(`이미지 다운로드 완료: ${outputPath}`);
}

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    "full-page": { type: "boolean", default: false },
    width: { type: "string", default: "1920" },
    height: { type: "string", default: "1080" },
    wait: { type: "string", default: "2" },
    device: { type: "string", default: "" },
  },
});

const [rawUrl, outputPath] = positionals;

if (!rawUrl || !outputPath) {
  console.error("Usage: node capture.mjs <URL> <저장경로> [옵션]");
  console.error("  --full-page        전체 페이지 스크롤 캡처");
  console.error("  --width <number>   뷰포트 너비 (기본: 1920)");
  console.error("  --height <number>  뷰포트 높이 (기본: 1080)");
  console.error("  --wait <number>    로드 후 대기 시간(초) (기본: 2)");
  console.error('  --device <name>    모바일 디바이스 에뮬레이션 (예: "iPhone 14")');
  process.exit(1);
}

const url =
  rawUrl.startsWith("http://") || rawUrl.startsWith("https://")
    ? rawUrl
    : `https://${rawUrl}`;

const fullPage = values["full-page"];
const width = parseInt(values.width, 10);
const height = parseInt(values.height, 10);
const waitSec = parseFloat(values.wait);
const deviceName = values.device;

const absOutput = resolve(outputPath);
mkdirSync(dirname(absOutput), { recursive: true });

// 이미지 URL이면 직접 다운로드하고 종료
if (await isImageUrl(url)) {
  try {
    await downloadImage(url, absOutput);
    process.exit(0);
  } catch (e) {
    console.error(e.message);
    process.exit(1);
  }
}

// 웹사이트 URL이면 Playwright로 스크린샷 캡처
const browser = await chromium.launch({ headless: true });

let context;
if (deviceName) {
  const deviceConfig = devices[deviceName];
  if (!deviceConfig) {
    const similar = Object.keys(devices)
      .filter((d) => d.toLowerCase().includes(deviceName.toLowerCase()))
      .slice(0, 5);
    console.error(`'${deviceName}'를 찾을 수 없습니다.`);
    if (similar.length) {
      console.error(`유사한 디바이스: ${similar.join(", ")}`);
    }
    await browser.close();
    process.exit(1);
  }
  context = await browser.newContext({ ...deviceConfig });
} else {
  context = await browser.newContext({
    viewport: { width, height },
    deviceScaleFactor: 1,
  });
}

const page = await context.newPage();

try {
  await page.goto(url, { waitUntil: "networkidle", timeout: 30000 });
} catch (e) {
  console.error(`networkidle failed, retrying with load: ${e.message}`);
  try {
    await page.goto(url, { waitUntil: "load", timeout: 30000 });
  } catch (e) {
    console.error(`페이지 로드 실패: ${e.message}`);
    await browser.close();
    process.exit(1);
  }
}

if (waitSec > 0) {
  await page.waitForTimeout(waitSec * 1000);
}

await page.screenshot({ path: absOutput, fullPage, type: "png" });
await browser.close();

console.log(`캡처 완료: ${absOutput}`);
