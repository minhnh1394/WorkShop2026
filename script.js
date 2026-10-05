/**
 * Tạo bảng yêu cầu đề bài MOS (Word / PowerPoint).
 * Nội dung được mã hóa (encodeURIComponent) để truyền qua tham số URL sang mos_notice.html.
 */
function buildMOSRequest(cfg) {
    const rows = cfg.questions.map((q, i) => `
        <tr>
            <td style="text-align:center; padding:10px; border:1px solid #ddd;">${i + 1}</td>
            <td style="padding:10px; border:1px solid #ddd;">${q}</td>
            <td style="text-align:center; padding:10px; border:1px solid #ddd;">200</td>
        </tr>`).join('');

    return encodeURIComponent(`
        <h3 style="text-align:center;">${cfg.title}</h3>
        <p style="text-align:center; font-weight:600; color:#0056b3;">TRẢI NGHIỆM – PROJECT A</p>
        <table style="width:100%; border-collapse:collapse; margin:15px 0;">
            <thead>
                <tr style="background:#0056b3; color:#fff;">
                    <th style="padding:10px; border:1px solid #ddd; width:60px;">Câu</th>
                    <th style="padding:10px; border:1px solid #ddd;">Yêu cầu</th>
                    <th style="padding:10px; border:1px solid #ddd; width:80px;">Điểm</th>
                </tr>
            </thead>
            <tbody>
                ${rows}
                <tr style="font-weight:bold; background:#f5f7fa;">
                    <td style="padding:10px; border:1px solid #ddd;"></td>
                    <td style="padding:10px; border:1px solid #ddd; text-align:right;">Tổng điểm</td>
                    <td style="text-align:center; padding:10px; border:1px solid #ddd;">1000</td>
                </tr>
            </tbody>
        </table>

        <div style="margin-top:20px;">
            <p style="font-weight:bold; color:#d9534f;">Lưu và chấm điểm:</p>
            <ol style="color:#333;">
                <li>Làm bài trên file <b>${cfg.workFile}</b> vừa được tải xuống.</li>
                <li>Lưu bài với tên theo cấu trúc: <b>${cfg.nameFormat}</b></li>
                <li>Tải file chấm điểm <a href="${cfg.gradeFile}" style="color:#007bff; text-decoration:underline;">tại đây</a>,
                    mở file (nhấn <b>Enable Content</b> để bật Macro), nhấn nút <b>Chấm điểm</b> rồi chọn file bài làm của bạn.</li>
                <li>Có thể xem lại đề bài dạng Word <a href="${cfg.requestFile}" style="color:#007bff; text-decoration:underline;">tại đây</a>.</li>
            </ol>
        </div>
    `);
}

const MOS_WORD_REQUEST = buildMOSRequest({
    title: 'ĐỀ KIỂM TRA MICROSOFT WORD',
    workFile: 'Word_Project_A_BaiLam.docx',
    nameFormat: 'HoTen_BoPhan_ChucDanh.docx',
    gradeFile: 'MOS_FILES/Word/Word_Project_A_ChamDiem.xlsm',
    requestFile: 'MOS_FILES/Word/Word_Project_A_YeuCau.docx',
    questions: [
        `Đổi lề trang của toàn bộ tài liệu (đơn vị inch):
            <ul>
                <li>Lề trên và dưới: <b>0.8</b></li>
                <li>Lề trái và phải: <b>1</b></li>
            </ul>`,
        `Ngắt trang kiểu <b>Next Page</b> vào trước tiêu đề “<b>About IIG Academy</b>”.`,
        `Trong trang 1, đặt số thứ tự cho “<b>English Tests</b>” để tiếp tục với số thứ tự là <b>2</b>.`,
        `Trong phần thuộc tính (properties) của tệp hãy thêm các nội dung:
            <ul>
                <li>“<b>IIG Academy</b>” vào danh mục (category)</li>
                <li>“<b>MOS Word 2019</b>” vào tiêu đề (title).</li>
            </ul>`,
        `Trong đoạn văn bản “<b>Contact us: IIG Viet Nam</b>” hãy tạo liên kết đến địa chỉ website: <b>https://iigvietnam.com/</b>`
    ]
});

const MOS_PPT_REQUEST = buildMOSRequest({
    title: 'ĐỀ KIỂM TRA MICROSOFT POWERPOINT',
    workFile: 'PPoint_Project_A_BaiLam.pptx',
    nameFormat: 'Lop_HoTen_NgaySinh.pptx',
    gradeFile: 'MOS_FILES/PowerPoint/PPoint_Project_A_ChamDiem.xlsm',
    requestFile: 'MOS_FILES/PowerPoint/PowerPoint_Project_A_YeuCau.docx',
    questions: [
        `Trong trang chiếu số 4 (<b>Most Popular</b>), áp dụng định dạng cho phần tiêu đề giống với các trang chiếu khác.`,
        `Trong trang chiếu số 2 (<b>Adventurous Trails!</b>), áp dụng hiệu ứng cho hình ảnh (Picture Effects) là <b>Soft Edges - 5 Point</b>.`,
        `Đổi tên Section “<b>Part 2</b>” thành “<b>Other</b>”.`,
        `Trong trang chiếu số 3 (<b>Why Mountain Bike Moab, Utah?</b>), thêm nhãn số liệu (Data Label) vào biểu đồ, vị trí đặt ở phía trên bên ngoài đỉnh cột.`,
        `Trong trang chiếu số <b>1</b>, hãy đặt độ trong suốt cho hình nền (Background) là <b>0%</b>.`
    ]
});

/**
 * Nội dung thông báo chung cho Excel.
 */
const MOS_GENERAL_NOTICE = (software) => encodeURIComponent(`
    <i class="fas fa-exclamation-triangle notice-icon" style="font-size: 3em; color: #ffc107;"></i>
    <h2>Chuẩn Bị Thực Hành MOS ${software}</h2>
    <p>File đề thi đã được tải xuống máy tính của bạn.</p>
`);


/**
 * Hàm xử lý bài thi MOS (Chuyển hướng đến trang thông báo và kích hoạt tải file bài làm)
 * @param {string} software - Tên phần mềm (Excel, Word, PowerPoint)
 */
function startMOS(software) {
    let downloadURL = '';
    let requestContent = '';

    if (software === 'Excel') {
        downloadURL = 'MOS_FILES/Excel/MOS_Excel_01.xlsm';
        requestContent = MOS_GENERAL_NOTICE('Excel');
    } else if (software === 'Word') {
        downloadURL = 'MOS_FILES/Word/Word_Project_A_BaiLam.docx';
        requestContent = MOS_WORD_REQUEST;
    } else if (software === 'PowerPoint') {
        downloadURL = 'MOS_FILES/PowerPoint/PPoint_Project_A_BaiLam.pptx';
        requestContent = MOS_PPT_REQUEST;
    } else {
        alert('Lỗi: Phần mềm MOS không hợp lệ.');
        return;
    }

    // Chuyển hướng đến trang thông báo, kèm link tải file VÀ nội dung yêu cầu
    window.location.href = `mos_notice.html?file=${encodeURIComponent(downloadURL)}&content=${requestContent}`;
}


/**
 * Hàm chọn ngẫu nhiên một bài thi và chuyển hướng cho các môn trắc nghiệm (IC3, GenAI, NLS).
 * Sử dụng cấu trúc thư mục [Môn Học]/[Tên Thư Mục Bài Thi]/index.html
 */
function startTest(subject, level) {
    let basePath = '';
    let folders = [];
    let displayName = '';
    const numTests = 5; // Số lượng bài thi (từ 01 đến 05)

    // --- 1. XÁC ĐỊNH DANH SÁCH THƯ MỤC CÓ THỂ CÓ ---
    
    if (subject === 'gen_ai') {
        basePath = 'GENAI';
        for (let i = 1; i <= numTests; i++) {
            folders.push(`GENAI_${String(i).padStart(2, '0')} (Published)`);
        }
        displayName = 'GENAI (AI TẠO SINH)';

    } else if (subject === 'nang_luc_so') {
        basePath = 'NLS';
        for (let i = 1; i <= numTests; i++) {
            folders.push(`NangLucSo_${String(i).padStart(2, '0')} (Published)`);
        }
        displayName = 'NĂNG LỰC SỐ';

    } else if (subject === 'ic3_gs6' || subject === 'ic3_gs6_spark') {
        if (!level) {
            alert('Lỗi: Cấp độ chưa được xác định cho bài thi này!');
            return;
        }

        const levelName = level.toUpperCase().replace('_', ' ');
        basePath = (subject === 'ic3_gs6') ? 'IC3 GS6' : 'IC3 GS6 Spark';
        
        displayName = `${basePath.toUpperCase()} - ${levelName}`;

        for (let i = 1; i <= numTests; i++) {
            folders.push(`${basePath} ${levelName}_${String(i).padStart(2, '0')} (Published)`);
        }

    } else {
        alert('Lỗi: Môn học không hợp lệ!');
        return;
    }

    // --- 2. CHỌN NGẪU NHIÊN VÀ CHUYỂN HƯỚNG ---
    
    if (folders.length === 0) {
        alert('Lỗi: Không tìm thấy bài thi nào cho môn này. Vui lòng kiểm tra lại tên thư mục.');
        return;
    }

    // Chọn ngẫu nhiên một thư mục trong danh sách folders
    const randomIndex = Math.floor(Math.random() * folders.length);
    const selectedFolder = folders[randomIndex];

    // Đường dẫn cuối cùng: [Base Path]/[Selected Folder]/index.html
    const finalURL = `${basePath}/${selectedFolder}/index.html`;

    // 4. Thông báo và chuyển hướng người dùng
    alert(`Đang mở bài thi thử ${displayName} (Mã bài: ${selectedFolder}). Chúc bạn làm bài tốt!`);
    
    window.location.href = finalURL;

}
