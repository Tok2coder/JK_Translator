// server.js
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

// 템플릿 파일 경로 설정
const templatesPath = path.join(__dirname, 'templates');

// 템플릿 폴더가 없으면 생성
if (!fs.existsSync(templatesPath)) {
    fs.mkdirSync(templatesPath, { recursive: true });
    console.log('템플릿 폴더가 생성되었습니다:', templatesPath);
}

// 템플릿에 번역 값 삽입 API
app.post('/api/fill-template', express.json(), async (req, res) => {
    try {
        // 요청 본문에서 템플릿 파일명과 치환 값 추출
        const { templateName, replacements } = req.body;
        
        if (!templateName || !replacements) {
            return res.status(400).json({ error: '템플릿 이름과 치환 값이 필요합니다.' });
        }
        
        // 템플릿 파일 경로
        const templatePath = path.join(templatesPath, templateName);
        
        // 파일 존재 여부 확인
        if (!fs.existsSync(templatePath)) {
            return res.status(404).json({ error: `템플릿 파일을 찾을 수 없습니다: ${templateName}` });
        }
        
        // 템플릿 파일 읽기
        const templateBuffer = fs.readFileSync(templatePath);
        
        // JSZip 초기화 및 파일 로드
        const JSZip = require('jszip');
        const zip = new JSZip();
        await zip.loadAsync(templateBuffer);
        
        // document.xml 파일 찾기
        const documentXml = await zip.file("word/document.xml").async("string");
        
        // 로그 - 처리 전 document.xml 내용의 일부 출력
        console.log("치환 전 XML 샘플:", documentXml.substring(0, 500) + "...");
        
        // 텍스트 치환하기 - 간단한 문자열 치환 방식 사용
        let modifiedDocumentXml = documentXml;
        
        for (const [pattern, replacement] of Object.entries(replacements)) {
            // XML 이스케이프 처리
            const escapedReplacement = String(replacement)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
            
            // 단순 문자열 치환 (정규표현식 없이)
            modifiedDocumentXml = modifiedDocumentXml.split(pattern).join(escapedReplacement);
            
            // 로그 - 각 패턴 치환 과정
            console.log(`'${pattern}' -> '${escapedReplacement}'`);
        }
        
        // 로그 - 처리 후 XML 내용에 패턴이 남아있는지 확인
        for (const pattern of Object.keys(replacements)) {
            if (modifiedDocumentXml.includes(pattern)) {
                console.log(`경고: 패턴 '${pattern}'이 여전히 XML에 존재함`);
            }
        }
        
        // 수정된 XML을 다시 압축 파일에 넣기
        zip.file("word/document.xml", modifiedDocumentXml);
        
        // DOCX 파일로 생성
        const content = await zip.generateAsync({ type: "nodebuffer" });
        
        // 응답 설정
        res.setHeader('Content-Disposition', `attachment; filename=translated_${templateName}`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(content);
    } catch (error) {
        console.error('템플릿 처리 오류:', error);
        res.status(500).json({ error: '템플릿 처리 중 오류가 발생했습니다: ' + error.message });
    }
});

app.use(express.static('public'));
app.use(express.static('./'));  // 현재 디렉토리의 정적 파일 제공
// 루트 경로에서 JK_translator.html 제공
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'JK_translator.html'));
  });


// DeepL API 키 (실제 사용 시에는 환경 변수로 관리)
// 여기에 DeepL API 키를 입력하세요
const DEEPL_API_KEY = '580a6175-23d0-4fce-a577-4cd7b6da73fe:fx';

// 도로명주소 API Key (실제 사용 시에는 환경 변수로 관리)
// 여기에 도로명주소 API Key를 입력하세요
const JUSO_API_KEY = 'devU01TX0FVVEgyMDI1MDQwMzAwMTgzMzExNTYwNjc=';

// 미들웨어 설정
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));
app.use(express.static('./'));  // 현재 디렉토리의 정적 파일 제공

// 도로명주소 영문변환 API 호출 함수
async function translateAddressToEnglish(koreanAddress) {
    try {
        console.log('도로명주소 변환 시작:', koreanAddress);
        
        // 출생장소 + 신고일 + 신고인 패턴 확인
        if (koreanAddress.includes('[출생장소]') && 
            (koreanAddress.includes('[신고일]') || koreanAddress.includes('[신고인]'))) {
            
            // 출생장소 추출
            const birthPlaceMatch = koreanAddress.match(/\[출생장소\]\s*([^\[]+)/);
            if (!birthPlaceMatch) return `Registration Address: ${koreanAddress}`;
            
            let birthPlace = birthPlaceMatch[1].trim();
            // [신고일] 전까지만 주소로 변환
            const reportDateIndex = birthPlace.indexOf('[신고일]');
            if (reportDateIndex !== -1) {
                birthPlace = birthPlace.substring(0, reportDateIndex).trim();
            }
            
            // 주소만 API로 변환
            let birthPlaceEng = '';
            try {
                // 도로명주소 API 호출 (주소 부분만)
                const encodedAddress = encodeURIComponent(birthPlace);
                const apiUrl = `https://business.juso.go.kr/addrlink/addrEngApi.do?currentPage=1&countPerPage=10&keyword=${encodedAddress}&confmKey=${JUSO_API_KEY}&resultType=json`;
                console.log('도로명주소 API 호출 (출생장소):', apiUrl);
                
                const response = await axios.get(apiUrl);
                
                if (response.data && 
                    response.data.results && 
                    response.data.results.common.totalCount > 0 &&
                    response.data.results.juso &&
                    response.data.results.juso[0]) {
                    
                    birthPlaceEng = response.data.results.juso[0].jibunAddr;
                    console.log('변환된 출생장소:', birthPlaceEng);
                } else {
                    birthPlaceEng = birthPlace;
                    console.log('출생장소 API 결과 없음, 원본 유지');
                }
            } catch (e) {
                console.error('출생장소 주소 변환 오류:', e);
                birthPlaceEng = birthPlace;
            }
            
            // 신고일 처리
            let reportDatePart = '';
            const reportDateMatch = koreanAddress.match(/\[신고일\]\s*(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)/);
            if (reportDateMatch) {
                const koreanDate = reportDateMatch[1].trim();
                const dateMatch = koreanDate.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
                if (dateMatch) {
                    const [_, year, month, day] = dateMatch;
                    reportDatePart = `[Report Date] ${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                } else {
                    reportDatePart = `[Report Date] ${koreanDate}`;
                }
            }
            
            // 신고인 처리
            let reporterPart = '';
            const reporterMatch = koreanAddress.match(/\[신고인\]\s*([^\[]+)/);
            if (reporterMatch) {
                const reporter = reporterMatch[1].trim();
                if (reporter === '부') {
                    reporterPart = '[Declarant] Father';
                } else if (reporter === '모') {
                    reporterPart = '[Declarant] Mother';
                } else {
                    reporterPart = `[Declarant] ${reporter}`;
                }
            }
            
            // 최종 결과 조합
            let result = `[Birth Place] ${birthPlaceEng}`;
            if (reportDatePart) {
                result += ` ${reportDatePart}`;
            }
            if (reporterPart) {
                result += ` ${reporterPart}`;
            }
            
            console.log('출생장소/신고일/신고인 최종 변환 결과:', result);
            return result;
        }
        
        // 일반 주소 처리 (기존 코드)
        let cleanAddress = koreanAddress;
        let prefixMatch = null;
        
        if (koreanAddress.includes('[')) {
            prefixMatch = koreanAddress.match(/\[(.*?)\]\s*(.*)/);
            if (prefixMatch) {
                cleanAddress = prefixMatch[2]; // 접두어 제외한 순수 주소 부분
                console.log('접두어 제거된 주소:', cleanAddress);
                
                // 첫 번째 접두어까지만 고려 (신고일, 신고인 등은 별도 처리)
                const endOfAddress = cleanAddress.indexOf('[');
                if (endOfAddress !== -1) {
                    cleanAddress = cleanAddress.substring(0, endOfAddress).trim();
                    console.log('추가 접두어 제거된 주소:', cleanAddress);
                }
            }
        }
        
        // URL 인코딩
        const encodedAddress = encodeURIComponent(cleanAddress);
        
        // API 호출
        const apiUrl = `https://business.juso.go.kr/addrlink/addrEngApi.do?currentPage=1&countPerPage=10&keyword=${encodedAddress}&confmKey=${JUSO_API_KEY}&resultType=json`;
        
        console.log('도로명주소 API 호출:', apiUrl);
        
        const response = await axios.get(apiUrl);
        
        // 응답 확인 및 데이터 추출
        if (response.data && 
            response.data.results && 
            response.data.results.common.totalCount > 0 &&
            response.data.results.juso &&
            response.data.results.juso[0]) {
            
            const jusoResult = response.data.results;
            
            // 에러 코드 확인
            if (jusoResult.common.errorCode !== "0") {
                console.error('주소 변환 API 오류:', jusoResult.common.errorMessage);
                return `Registration Address: ${koreanAddress}`;
            }
            
            // 결과 중 첫 번째 항목 사용
            const firstResult = jusoResult.juso[0];
            
            // 영문 주소 구성 - jibunAddr 필드 사용
            let engAddress = firstResult.jibunAddr;
            console.log('변환된 영문 주소:', engAddress);
            
            // 접두어가 있었던 경우 영문 접두어 추가
            if (prefixMatch) {
                const prefix = prefixMatch[1];
                let engPrefix = prefix;
                
                // 접두어 영문 변환
                if (prefix === '출생장소') engPrefix = 'Birth Place';
                if (prefix === '등록기준지') engPrefix = 'Registration Address';
                
                const result = `[${engPrefix}] ${engAddress}`;
                console.log('최종 영문 주소 (접두어 포함):', result);
                return result;
            }
            
            console.log('최종 영문 주소:', engAddress);
            return engAddress;
        } else {
            console.log('도로명주소 API 결과 없음:', koreanAddress);
            // API 결과가 없을 경우 기본 영문 표기 추가
            if (koreanAddress.startsWith('[')) {
                return koreanAddress; // 이미 처리된 주소는 그대로 반환
            }
            return `Registration Address: ${koreanAddress}`;
        }
    } catch (error) {
        console.error('도로명주소 API 호출 오류:', error.message);
        // 오류 발생 시 기본 영문 표기 추가
        return `Registration Address: ${koreanAddress}`;
    }
}

// 특수 형식 주소 처리 (출생장소 + 신고일 + 신고인)
async function translateSpecialAddress(text) {
    try {
        // 1. 출생장소 추출 및 변환
        const birthPlaceMatch = text.match(/\[출생장소\]\s*([^\[]+)/);
        if (!birthPlaceMatch) return text;
        
        const birthPlaceWithExtra = birthPlaceMatch[1].trim();
        
        // [신고일] 전까지만 주소로 간주
        let birthPlace = birthPlaceWithExtra;
        const reportDateIndex = birthPlaceWithExtra.indexOf('[신고일]');
        
        if (reportDateIndex !== -1) {
            birthPlace = birthPlaceWithExtra.substring(0, reportDateIndex).trim();
        }
        
        console.log('추출된 출생장소:', birthPlace);
        
        // 주소 변환
        let translatedBirthPlace;
        try {
            translatedBirthPlace = await translateAddressToEnglish(birthPlace);
            // 접두어가 추가되어 있을 수 있으므로 제거
            if (translatedBirthPlace.startsWith('Birth Place:') || 
                translatedBirthPlace.startsWith('[Birth Place]')) {
                translatedBirthPlace = translatedBirthPlace.replace(/^\[Birth Place\]:|^Birth Place:/, '').trim();
            }
        } catch (e) {
            console.error('출생장소 변환 오류:', e);
            translatedBirthPlace = birthPlace;
        }
        
        // 2. 신고일 추출 및 변환
        let reportDatePart = '';
        const reportDateMatch = text.match(/\[신고일\]\s*(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)/);
        
        if (reportDateMatch) {
            const koreanDate = reportDateMatch[1].trim();
            const dateMatch = koreanDate.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
            
            if (dateMatch) {
                const [_, year, month, day] = dateMatch;
                reportDatePart = `[Report Date] ${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
            } else {
                reportDatePart = `[Report Date] ${koreanDate}`;
            }
        }
        
        // 3. 신고인 추출 및 변환
        let reporterPart = '';
        const reporterMatch = text.match(/\[신고인\]\s*([^\[]+)(?:\s*$|(?=\s*\[))/);
        
        if (reporterMatch) {
            const reporter = reporterMatch[1].trim();
            if (reporter === '부') {
                reporterPart = '[Declarant] Father';
            } else if (reporter === '모') {
                reporterPart = '[Declarant] Mother';
            } else {
                reporterPart = `[Declarant] ${reporter}`;
            }
        }
        
        // 4. 결과 조합
        let result = `[Birth Place] ${translatedBirthPlace}`;
        if (reportDatePart) {
            result += ` ${reportDatePart}`;
        }
        if (reporterPart) {
            result += ` ${reporterPart}`;
        }
        
        console.log('최종 변환 결과:', result);
        return result;
    } catch (error) {
        console.error('특수 주소 처리 오류:', error);
        return text;
    }
}

// 텍스트 번역 API (DeepL 사용)
app.post('/api/translate', async (req, res) => {
    console.log('번역 요청 받음');
    
    try {
        const { values, documentType, customWords, forUsa } = req.body;
        
        if (!values) {
            return res.status(400).json({ error: '유효하지 않은 요청입니다. values가 필요합니다.' });
        }
        
        console.log('번역 요청 데이터 타입:', {
            title: typeof values.title,
            registrationNumber: typeof values.registrationNumber,
            table1: Array.isArray(values.table1) ? `Array(${values.table1.length})` : typeof values.table1,
            table2: Array.isArray(values.table2) ? `Array(${values.table2.length})` : typeof values.table2,
            table3: Array.isArray(values.table3) ? `Array(${values.table3.length})` : typeof values.table3
        });

        // 번역할 필드 준비
        const translations = {};
        
        // 기본 정보 번역
        if (values.title) {
            if (values.title === '기본증명서 (상세)') {
                translations.title = 'Basic Certificate (Detailed)';
            } else if (values.title === '가족관계증명서') {
                translations.title = 'Family Relations Certificate';
            } else if (values.title === '혼인관계증명서') {
                translations.title = 'Marriage Certificate';
            } else {
                translations.title = values.title;
            }
        }
        
        // 등록기준지 영문 변환 (도로명주소 API 사용)
        if (values.registrationNumber) {
            try {
                console.log('등록기준지 변환 시작:', values.registrationNumber);
                translations.registrationNumber = await translateAddressToEnglish(values.registrationNumber);
                console.log('등록기준지 영문 변환 완료:', translations.registrationNumber);
            } catch (addressError) {
                console.error('등록기준지 변환 오류:', addressError);
                translations.registrationNumber = `Registration Address: ${values.registrationNumber}`;
            }
        }
        
        // 테이블 데이터 번역 초기화
        translations.table1 = [];
        translations.table2 = [];
        translations.table3 = [];

        // DeepL API 키가 유효하면 DeepL API 사용
        if (DEEPL_API_KEY && DEEPL_API_KEY !== 'your-deepl-api-key') {
            try {
                // 번역할 텍스트 준비
                const textsToTranslate = [];
                const textMapping = []; // 원본 텍스트의 위치 정보 저장
                
                // table1 번역 텍스트 준비 (작성정보)
                if (values.table1 && Array.isArray(values.table1)) {
                    values.table1.forEach((row, rowIndex) => {
                        if (Array.isArray(row)) {
                            row.forEach((cell, colIndex) => {
                                if (cell && cell.trim() !== '') {
                                    // 특정단어 처리
                                    let processedText = cell;
                                    for (const [koreanWord, englishWord] of Object.entries(customWords || {})) {
                                        if (cell.includes(koreanWord)) {
                                            processedText = processedText.replace(new RegExp(koreanWord, 'g'), `<keep>${englishWord}</keep>`);
                                        }
                                    }
                                    
                                    textsToTranslate.push(processedText);
                                    textMapping.push({ type: 'table1', row: rowIndex, col: colIndex });
                                }
                            });
                        }
                    });
                    
                    // table1 초기화
                    translations.table1 = values.table1.map(row => Array(row.length).fill(''));
                }
                
                // table2 번역 텍스트 준비 (인적사항)
                if (values.table2 && Array.isArray(values.table2)) {
                    values.table2.forEach((row, rowIndex) => {
                        if (Array.isArray(row)) {
                            row.forEach((cell, colIndex) => {
                                if (cell && cell.trim() !== '') {
                                    // 특정단어 처리
                                    let processedText = cell;
                                    for (const [koreanWord, englishWord] of Object.entries(customWords || {})) {
                                        if (cell.includes(koreanWord)) {
                                            processedText = processedText.replace(new RegExp(koreanWord, 'g'), `<keep>${englishWord}</keep>`);
                                        }
                                    }
                                    
                                    textsToTranslate.push(processedText);
                                    textMapping.push({ type: 'table2', row: rowIndex, col: colIndex });
                                }
                            });
                        }
                    });
                    
                    // table2 초기화
                    translations.table2 = values.table2.map(row => Array(row.length).fill(''));
                }
                
                // table3 번역 텍스트 준비 (일반등록사항)
                if (values.table3 && Array.isArray(values.table3)) {
                    values.table3.forEach((row, rowIndex) => {
                        if (Array.isArray(row)) {
                            row.forEach((cell, colIndex) => {
                                if (cell && cell.trim() !== '') {
                                    // 출생장소나 주소 패턴인 경우 도로명주소 API 사용
                                    const isAddressLike = cell.includes('[출생장소]') || 
                                                       cell.includes('번지') || 
                                                       cell.includes('로 ') || 
                                                       cell.includes('길 ');
                                    
                                    if (isAddressLike) {
                                        // 도로명주소 API 호출 결과는 별도 처리 (번역 텍스트에 추가 안함)
                                        translateAddressToEnglish(cell).then(translatedAddress => {
                                            translations.table3[rowIndex][colIndex] = translatedAddress;
                                        }).catch(err => {
                                            console.error('주소 변환 오류:', err);
                                            translations.table3[rowIndex][colIndex] = cell;
                                        });
                                    } else {
                                        // 특정단어 처리
                                        let processedText = cell;
                                        for (const [koreanWord, englishWord] of Object.entries(customWords || {})) {
                                            if (cell.includes(koreanWord)) {
                                                processedText = processedText.replace(new RegExp(koreanWord, 'g'), `<keep>${englishWord}</keep>`);
                                            }
                                        }
                                        
                                        textsToTranslate.push(processedText);
                                        textMapping.push({ type: 'table3', row: rowIndex, col: colIndex });
                                    }
                                }
                            });
                        }
                    });
                    
                    // table3 초기화
                    translations.table3 = values.table3.map(row => Array(row.length).fill(''));
                }
                // 기타 필드 번역 텍스트 준비
                const otherFields = ['certText', 'issuanceDate', 'issuancePlace', 'issuanceTime', 'issuer', 'issuerContact', 'applicant', 'issueNumber'];
                
                otherFields.forEach(field => {
                    if (values[field] && values[field].trim() !== '') {
                        // 날짜 필드는 특별 처리
                        if (field === 'issuanceDate') {
                            const dateMatch = values[field].match(/(\d+)년\s*(\d+)월\s*(\d+)일/);
                            if (dateMatch) {
                                const [_, year, month, day] = dateMatch;
                                translations[field] = `${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                            } else {
                                textsToTranslate.push(values[field]);
                                textMapping.push({ type: 'field', field });
                            }
                        }
                        // 기본값이 미리 정의된 필드는 처리 생략
                        else if (field === 'certText') {
                            translations[field] = 'This is to certify that the above Basic Certificate (Detailed) is in accordance with the records of the Family Relationship Registry.';
                        }
                        // 나머지 필드는 번역
                        else {
                            textsToTranslate.push(values[field]);
                            textMapping.push({ type: 'field', field });
                        }
                    } else {
                        // 빈 값은 빈 문자열로 초기화
                        translations[field] = '';
                    }
                });
                
                // DeepL API 호출 (번역할 텍스트가 있는 경우에만)
                if (textsToTranslate.length > 0) {
                    console.log(`DeepL API 호출: ${textsToTranslate.length}개 텍스트 번역 중...`);
                    
                    const response = await axios.post(
                        'https://api-free.deepl.com/v2/translate',
                        {
                            text: textsToTranslate,
                            target_lang: 'EN',
                            preserve_formatting: true,
                            tag_handling: 'xml',
                            ignore_tags: ['keep']
                        },
                        {
                            headers: {
                                'Authorization': `DeepL-Auth-Key ${DEEPL_API_KEY}`,
                                'Content-Type': 'application/json'
                            }
                        }
                    );
                    
                    if (response.data && response.data.translations) {
                        const translatedTexts = response.data.translations.map(t => t.text);
                        
                        // 번역 결과를 매핑 정보를 통해 적절한 위치에 배치
                        translatedTexts.forEach((translatedText, index) => {
                            const mapping = textMapping[index];
                            
                            if (mapping.type === 'table1') {
                                translations.table1[mapping.row][mapping.col] = translatedText;
                            }
                            else if (mapping.type === 'table2') {
                                translations.table2[mapping.row][mapping.col] = translatedText;
                            }
                            else if (mapping.type === 'table3') {
                                translations.table3[mapping.row][mapping.col] = translatedText;
                            }
                            else if (mapping.type === 'field') {
                                translations[mapping.field] = translatedText;
                            }
                        });
                        
                        console.log('DeepL 번역 완료');
                    } else {
                        console.warn('DeepL API 응답 누락, 기본 번역 사용');
                        await applyBasicTranslation(values, translations, customWords);
                    }
                } else {
                    console.log('번역할 텍스트가 없음, 기본 번역 사용');
                    await applyBasicTranslation(values, translations, customWords);
                }
            } catch (deepLError) {
                console.error('DeepL API 오류:', deepLError.message);
                await applyBasicTranslation(values, translations, customWords);
            }
        } else {
            console.log('DeepL API 키가 설정되지 않음, 기본 번역 사용');
            await applyBasicTranslation(values, translations, customWords);
        }

        // 미국 제출용 옵션 처리 (기본증명서에 한함)
        if (documentType === 'basic' && forUsa) {
            if (translations.table3 && translations.table3[0] && translations.table3[0][1]) {
                translations.table3[0][1] += ' (Republic of Korea)';
            }
        }
        
        // 응답
        res.json({
            success: true,
            translations
        });
    } catch (error) {
        console.error('번역 처리 중 오류:', error);
        res.status(500).json({ 
            error: '번역 중 오류가 발생했습니다.',
            message: error.message
        });
    }
});

// 기본 번역 함수 (DeepL API 실패 시 사용)
async function applyBasicTranslation(values, translations, customWords) {
    console.log('기본 번역 적용 시작');
    
    // table1 기본 번역 (작성정보)
    if (values.table1 && Array.isArray(values.table1)) {
        translations.table1 = [];
        
        for (let i = 0; i < values.table1.length; i++) {
            const row = values.table1[i];
            const translatedRow = [];
            
            for (let j = 0; j < row.length; j++) {
                const cell = row[j];
                let translatedCell = "";
                
                // 주소 패턴 검사 (도로명주소 API 사용)
                if (cell && (cell.includes('번지') || cell.includes('로 ') || cell.includes('길 '))) {
                    try {
                        translatedCell = await translateAddressToEnglish(cell);
                    } catch (addrError) {
                        console.error('주소 변환 오류:', addrError);
                        translatedCell = cell;
                    }
                } else {
                    // 특정단어 처리
                    translatedCell = cell;
                    for (const [koreanWord, englishWord] of Object.entries(customWords || {})) {
                        if (cell && cell.includes(koreanWord)) {
                            translatedCell = translatedCell.replace(new RegExp(koreanWord, 'g'), englishWord);
                        }
                    }
                    
                    // 기본 번역
                    if (translatedCell === '작성') {
                        translatedCell = 'Creation';
                    } else if (translatedCell === '작성사유') {
                        translatedCell = 'Reason for Creation';
                    }
                    // 날짜 형식 변환 - 접두어 보존
                    else if (translatedCell && translatedCell.includes('가족관계등록부작성일')) {
                        const dateMatch = translatedCell.match(/(\d+)년\s*(\d+)월\s*(\d+)일/);
                        if (dateMatch) {
                            const [_, year, month, day] = dateMatch;
                            translatedCell = `[Family Relationship Register Creation Date] ${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                        }
                    }
                    // 작성사유 번역
                    else if (translatedCell && translatedCell.includes('작성사유')) {
                        if (translatedCell.includes('가족관계의 등록 등에 관한 법률 부칙 제3조제1항')) {
                            translatedCell = '[Reason for Creation] Article 3(1) of the Addenda to the Act on the Registration of Family Relations';
                        }
                    }
                    // 일반 날짜 형식 변환
                    else {
                        const dateMatch = translatedCell.match(/(\d+)년\s*(\d+)월\s*(\d+)일/);
                        if (dateMatch) {
                            const [_, year, month, day] = dateMatch;
                            translatedCell = `${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                        }
                    }
                }
                
                translatedRow.push(translatedCell);
            }
            
            translations.table1.push(translatedRow);
        }
    }

    // table2 기본 번역 (인적사항)
    if (values.table2 && Array.isArray(values.table2)) {
        translations.table2 = [];
        
        for (let i = 0; i < values.table2.length; i++) {
            const row = values.table2[i];
            const translatedRow = [];
            
            for (let j = 0; j < row.length; j++) {
                const cell = row[j];
                let translatedCell = "";
                
                // 주소 패턴 검사 (도로명주소 API 사용)
                if (cell && (cell.includes('번지') || cell.includes('로 ') || cell.includes('길 '))) {
                    try {
                        translatedCell = await translateAddressToEnglish(cell);
                    } catch (addrError) {
                        console.error('주소 변환 오류:', addrError);
                        translatedCell = cell;
                    }
                } else {
                    // 특정단어 처리
                    translatedCell = cell;
                    for (const [koreanWord, englishWord] of Object.entries(customWords || {})) {
                        if (cell && cell.includes(koreanWord)) {
                            translatedCell = translatedCell.replace(new RegExp(koreanWord, 'g'), englishWord);
                        }
                    }
                    
                    // 열 위치에 따른 번역
                    if (j === 0) { // 구분
                        if (translatedCell === '본인') {
                            translatedCell = 'Subject';
                        } else if (translatedCell === '부') {
                            translatedCell = 'Father';
                        } else if (translatedCell === '모') {
                            translatedCell = 'Mother';
                        }
                    } else if (j === 4) { // 성별
                        if (translatedCell === '남') {
                            translatedCell = 'Male';
                        } else if (translatedCell === '여') {
                            translatedCell = 'Female';
                        }
                    } else if (j === 2) { // 출생연월일
                        const dateMatch = translatedCell.match(/(\d+)년\s*(\d+)월\s*(\d+)일/);
                        if (dateMatch) {
                            const [_, year, month, day] = dateMatch;
                            translatedCell = `${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                        }
                    }
                }
                
                translatedRow.push(translatedCell);
            }
            
            translations.table2.push(translatedRow);
        }
    }
    
    // table3 기본 번역 (일반등록사항)
    if (values.table3 && Array.isArray(values.table3)) {
        translations.table3 = [];
        
        for (let i = 0; i < values.table3.length; i++) {
            const row = values.table3[i];
            const translatedRow = [];
            
            for (let j = 0; j < row.length; j++) {
                const cell = row[j];
                let translatedCell = "";
                
                // 특정단어 처리
                let processedCell = cell || '';
                for (const [koreanWord, englishWord] of Object.entries(customWords || {})) {
                    if (processedCell.includes(koreanWord)) {
                        processedCell = processedCell.replace(new RegExp(koreanWord, 'g'), englishWord);
                    }
                }
                
                // 열 위치에 따른 번역
                if (j === 0) { // 구분
                    if (processedCell === '출생') {
                        translatedCell = 'Birth';
                    } else {
                        translatedCell = processedCell;
                    }
                }
                // 출생 관련 정보 번역
                else if (processedCell.includes('[출생장소]')) {
                    // 1. 출생장소 부분 추출 및 처리
                    const birthPlaceMatch = processedCell.match(/\[출생장소\]\s*([^\[]+)/);
                    let birthPlace = '';
                    if (birthPlaceMatch) {
                        birthPlace = birthPlaceMatch[1].trim();
                        // [신고일] 이전까지만 추출
                        const reportDateIndex = birthPlace.indexOf('[신고일]');
                        if (reportDateIndex !== -1) {
                            birthPlace = birthPlace.substring(0, reportDateIndex).trim();
                        }
                    }
                    
                    // 2. 출생 장소 영문 변환
                    let engBirthPlace = '';
                    try {
                        if (birthPlace && birthPlace.includes('번지')) {
                            engBirthPlace = await translateAddressToEnglish(birthPlace);
                        } else {
                            engBirthPlace = birthPlace;
                        }
                    } catch (e) {
                        console.error('출생장소 주소 변환 오류:', e);
                        engBirthPlace = birthPlace;
                    }
                    
                    // 3. 신고일 추출 및 변환
                    let reportDatePart = '';
                    const reportDateMatch = processedCell.match(/\[신고일\]\s*(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)/);
                    if (reportDateMatch) {
                        const koreanDate = reportDateMatch[1].trim();
                        const dateMatch = koreanDate.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
                        if (dateMatch) {
                            const [_, year, month, day] = dateMatch;
                            reportDatePart = `[Report Date] ${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                        } else {
                            reportDatePart = `[Report Date] ${koreanDate}`;
                        }
                    }
                    
                    // 4. 신고인 추출 및 변환
                    let reporterPart = '';
                    const reporterMatch = processedCell.match(/\[신고인\]\s*([^\[]+)(?:\s*$|(?=\s*\[))/);
                    if (reporterMatch) {
                        const reporter = reporterMatch[1].trim();
                        if (reporter === '부') {
                            reporterPart = '[Declarant] Father';
                        } else if (reporter === '모') {
                            reporterPart = '[Declarant] Mother';
                        } else {
                            reporterPart = `[Declarant] ${reporter}`;
                        }
                    }
                    
                    // 5. 결과 조합
                    translatedCell = `[Birth Place] ${engBirthPlace}`;
                    if (reportDatePart) {
                        translatedCell += ` ${reportDatePart}`;
                    }
                    if (reporterPart) {
                        translatedCell += ` ${reporterPart}`;
                    }
                    
                    console.log('최종 변환 결과:', translatedCell);
                }
                // 주소 패턴 검사 (도로명주소 API 사용)
                else if (processedCell && (processedCell.includes('번지') || processedCell.includes('로 ') || processedCell.includes('길 '))) {
                    try {
                        translatedCell = await translateAddressToEnglish(processedCell);
                    } catch (addrError) {
                        console.error('주소 변환 오류:', addrError);
                        translatedCell = processedCell;
                    }
                }
                else {
                    // 접두어가 있는 텍스트 처리
                    const prefixMatch = processedCell.match(/\[(.*?)\](.*)/);
                    if (prefixMatch) {
                        const prefix = prefixMatch[1];
                        const content = prefixMatch[2].trim();
                        
                        // 접두어 번역
                        let translatedPrefix = prefix;
                        if (prefix === '출생장소') translatedPrefix = 'Birth Place';
                        if (prefix === '신고일') translatedPrefix = 'Report Date';
                        if (prefix === '신고인') translatedPrefix = 'Declarant';
                        
                        translatedCell = `[${translatedPrefix}] ${content}`;
                    }
                    // 일반 날짜 형식 변환
                    else {
                        const dateMatch = processedCell.match(/(\d+)년\s*(\d+)월\s*(\d+)일/);
                        if (dateMatch) {
                            const [_, year, month, day] = dateMatch;
                            translatedCell = `${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
                        } else {
                            translatedCell = processedCell;
                        }
                    }
                }
                
                translatedRow.push(translatedCell);
            }
            
            translations.table3.push(translatedRow);
        }
    }
    
    // 기타 정보 기본 번역
    translations.certText = 'This is to certify that the above Basic Certificate (Detailed) is in accordance with the records of the Family Relationship Registry.';

    // 발행 날짜 변환
    const issuanceDateMatch = values.issuanceDate ? values.issuanceDate.match(/(\d+)년\s*(\d+)월\s*(\d+)일/) : null;
    if (issuanceDateMatch) {
        const [_, year, month, day] = issuanceDateMatch;
        translations.issuanceDate = `${month.padStart(2, '0')}/${day.padStart(2, '0')}/${year}`;
    } else {
        translations.issuanceDate = values.issuanceDate || '';
    }
    
    // 나머지 기타 정보 번역
    translations.issuancePlace = values.issuancePlace ? `Issuing Authority: ${values.issuancePlace}` : '';
    translations.issuanceTime = values.issuanceTime ? `Issue Time: ${values.issuanceTime}` : '';
    translations.issuer = values.issuer || '';
    translations.issuerContact = values.issuerContact || '';
    translations.applicant = values.applicant ? `Applicant: ${values.applicant}` : '';
    translations.issueNumber = values.issueNumber ? `Issue Number: ${values.issueNumber}` : '';
    
    console.log('기본 번역 적용 완료');
}

// 서버 시작
app.listen(port, () => {
    console.log(`서버가 http://localhost:${port} 에서 실행 중입니다.`);
});

module.exports = app;