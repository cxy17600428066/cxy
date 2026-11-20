import fs from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
const ROOT='C:/Users/admin/Documents/New project';
const RUNTIME='C:/Program Files/WindowsApps/OpenAI.CodexPrimaryRuntime.v26-1119-907-0_26.1119.907.0_x64__3k8sg7r9htsxt/dependencies';
const SKILL='C:/Users/admin/.codex/plugins/cache/openai-primary-runtime/presentations/26.1119.11507/skills/presentations';
process.env.RUNTIME_NODE_MODULES=RUNTIME+'/node/node_modules';
const {finalizePresentation}=await import(pathToFileURL(SKILL+'/container_tools/artifact_tool_utils.mjs'));
const res=await finalizePresentation({
 workspaceDir:ROOT,
 candidatePath:ROOT+'/.japan-travel-build/candidate.pptx',
 finalPath:ROOT+'/日本5日游宣传/富士山与东京-杭州出发5日游.pptx',
 pythonExecutable:RUNTIME+'/python/python.exe',
 integrityValidatorPath:SKILL+'/container_tools/inspect_presentation_package_integrity.py',
 layoutValidatorPath:SKILL+'/container_tools/inspect_presentation_layout_geometry.py',
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit'],
 explicitTotalSlideCount:6,
 requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[],
 fontPolicy:{basis:'design',families:['Microsoft YaHei']},
 verifyArtifactToolImport:true,
 receiptPath:ROOT+'/.japan-travel-build/validation.json'
});
console.log(JSON.stringify({finalPath:res.finalPath,bytes:res.byteCount,integrity:res.packageIntegrity,layout:res.presentationLayout},null,2));
await fs.copyFile(ROOT+'/.japan-travel-build/renders/slide-1.png',ROOT+'/日本5日游宣传/旅行宣传封面.png');
