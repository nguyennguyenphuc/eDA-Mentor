---
name: goi-ngoai
description: Gọi model NGOÀI Claude — ChatGPT web (qua Chrome của người dùng, cổng debug 9222) và Gemini qua Antigravity CLI (agy). Dùng khi cần ý kiến thứ hai / người chấm khác model người viết, viết lại một đoạn văn, hoặc giao việc cơ khí (sửa script, thay chuỗi hàng loạt, đọc log tìm nguyên nhân, dựng .drawio) để tiết kiệm token Claude. Kích hoạt khi người dùng nói "hỏi ChatGPT", "gửi ChatGPT chấm", "nhờ Gemini/agy làm", "giao việc cơ khí", "chatgpt bridge", "agy".
---

# Gọi ngoài: ChatGPT web và agy (Gemini)

Một CLI: `python <skill>/scripts/goi_ngoai.py`. Claude **chỉ đọc tóm tắt ≤10 dòng** nó in ra; nguyên văn nằm
trong tệp. Mỗi lần gọi ghi một dòng vào `<skill>/nhat-ky/goi-ngoai.tsv` (để tính chi phí ngoài Claude).

```
python goi_ngoai.py --tu-kiem                                   # kiểm sẵn sàng, không gửi gì ra ngoài
python goi_ngoai.py chatgpt "<câu hỏi>" [--tep a.md ...] [--ra tra_loi.md]
python goi_ngoai.py chatgpt --prompt-tep cau_hoi.md --kieu cham --ra phieu.md
python goi_ngoai.py agy "<việc>" --tep thu_muc_con/a.py [--tep ...] [--suy-luan]
```

## Chọn đích

| Việc | Đích | Ghi chú |
|---|---|---|
| Ý kiến thứ hai, chấm/critic, đọc mù, viết lại đoạn văn | `chatgpt` | Model khác người viết → thoả luật "judge khác model". Temporary Chat mặc định |
| Phiếu chấm có khuôn | `chatgpt --kieu cham` | Prompt phải yêu cầu khối `###PHIEU###` … `###HET###`; thiếu khối thì coi là hỏng và hỏi lại |
| Việc code cơ khí: đổi tên, thay chuỗi, sửa script theo mô tả rõ | `agy` | Gemini 3.8 Flash |
| Việc cần suy luận: đọc log tìm nguyên nhân lỗi, dựng `.drawio` từ mô tả | `agy --suy-luan` | Gemini 3.1 Pro |
| Việc chỉ là chạy script | Bash trực tiếp | Không gọi ngoài |

## Mã thoát — hành xử theo mã

| Mã | Nghĩa | Claude làm gì |
|---|---|---|
| 0 | xong | đọc tóm tắt; với `chatgpt` mở tệp `--ra` nếu cần nội dung |
| 1 | agy chạy xong nhưng tự báo `LOI` | đọc dòng tóm tắt, sửa đề bài hoặc tự làm |
| 2 | FALLBACK: hỏng 3 lần | tự làm việc đó (hoặc spawn một subagent), ghi lý do |
| 3 | CHƯA-CÀI: thiếu agy / thiếu Python có `websocket-client` | tự làm bằng Bash; báo người dùng cài |

## Chuẩn bị một lần (người dùng tự làm)

**ChatGPT web** — script không tự mở Chrome:
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%LOCALAPPDATA%\chatgpt-bridge-profile"
```
Đăng nhập chatgpt.com trong cửa sổ đó và để tab mở. Bắt buộc `--user-data-dir` riêng (Chrome ≥136 từ chối cổng
debug trên profile mặc định). Kiểm: `python <skill>/scripts/chatgpt_bridge.py --selftest` (thêm `--live` để gửi
thử một lượt thật). Python chạy cầu nối cần `websocket-client` + `requests`; chọn bằng biến `CHATGPT_PYTHON`
nếu `python` mặc định thiếu.

**agy** — Antigravity CLI cài sẵn tại `%LOCALAPPDATA%\agy\bin\agy.exe` (agy 1.2.14). Lệnh và model ở
`<skill>/cau-hinh/agy-cmd.json`; đổi đường dẫn bằng biến `AGY_EXE`; tên model lấy từ `agy models`.

## An toàn (bắt buộc)

- agy chạy với `--dangerously-skip-permissions`: Gemini **tự chạy mọi lệnh và sửa mọi tệp trong thư mục làm việc**
  (= thư mục chung của các `--tep`). Chỉ truyền tệp nằm trong **một thư mục con** của việc. Không bao giờ truyền
  tệp ở gốc dự án, trong `.claude/`, hay tệp chứa mật khẩu/cookie.
- Sau khi agy báo `XONG`: kiểm lại tệp đã sửa (diff/đọc/chạy) — agy tự khai kết quả, không phải bằng chứng.
- ChatGPT chỉ thấy đúng những gì được dán vào prompt. Không dán đáp án/khoá chấm vào cùng prompt với bài cần
  chấm; không dán dữ liệu nhạy cảm (mật khẩu, thông tin cá nhân học viên).

## Đường cụt đã thử — đừng thử lại

Chi tiết ở docstring đầu `scripts/chatgpt_bridge.py`: gọi `backend-api/conversation` (cần proof-of-work token của
trang), gán `.innerText` cho composer (ProseMirror bỏ qua), `InputEvent` tổng hợp (isTrusted=false),
`time.sleep` cố định để đợi trả lời (cắt cụt câu trả lời), MCP/tunnel. Script phụ thuộc DOM chatgpt.com:
khi OpenAI đổi giao diện, chạy `--selftest` để xem selector nào còn sống, sửa bảng `SEL` ở đầu tệp.

## Nguồn gốc

Rút từ `G:\My Drive\09._PC\Giang-day\Soan bai\Project DA\skill\.claude\skills\academic-writing-v3\scripts\`
(`chatgpt_bridge.py` bản 2026-10-02 + phần `co-khi` của `giao_viec.py`). Bản gốc giữ nguyên. Thêm `--kieu hoi`
(hỏi đáp tự do) cho cầu nối; bỏ các vai chấm bài giảng LaTeX và tầng Codex.
