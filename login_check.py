import asyncio
import os
import sys
import re
import csv
import datetime
from playwright.async_api import async_playwright

# ================= 配置区 =================
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
USER_DATA_PATH = os.path.join(os.environ['LOCALAPPDATA'], r"Microsoft\Edge\User Data\Automation")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
REPORT_DIR = os.path.join(os.getcwd(), "reports")
TASK_FILE = os.path.join(os.getcwd(), "tasks.txt")
# ==========================================

# 确保必要的目录存在
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)

async def auto_clean_modals(page):
    """清理弹窗干扰"""
    try:
        await page.evaluate("""() => {
            const selectors = ['.ant-modal-root', '.ant-modal-mask', '.xm-notification', '.ant-modal-wrap'];
            selectors.forEach(s => {
                const el = document.querySelector(s);
                if(el) el.remove();
            });
            document.body.style.overflow = 'auto';
        }""")
    except: pass

async def check_login_status(page):
    """【容灾机制】登录哨兵：检测登录态并支持人工介入"""
    print(f"[*] 正在执行登录态健康检查...")
    await page.goto("https://crm.xiaoman.cn/", wait_until="commit")
    await asyncio.sleep(4)
    
    current_url = page.url
    if "login" in current_url.lower():
        print("\n" + "="*50)
        print(" [⚠️ 警告] 检测到 CRM 登录态已失效！")
        print(" 请在弹出的浏览器窗口中手动完成登录（密码或扫码）。")
        print(" 登录成功并看到【主页/工作台】后，在此按下回车键继续。")
        print("="*50 + "\n")
        input(">>> 确认已登录后，请按【Enter】键继续：")
        print("[*] 收到继续指令，重新校验状态...")
    else:
        print("[+] 登录态有效，准备进入工作流。")

async def process_single_task(context, page, product_id):
    """处理单个产品的核心流转逻辑"""
    print(f"\n[{product_id}] 开始处理...")
    try:
        # 1. 搜索与跳转
        await page.goto("https://crm.xiaoman.cn/product", wait_until="commit")
        await asyncio.sleep(2)
        await auto_clean_modals(page)

        search_input = page.locator(".ow-fixed-render-name-inner", has_text="编号/型号/名称").first
        await search_input.click()
        await page.keyboard.press("Control+A")
        await page.keyboard.press("Backspace")
        await page.keyboard.type(product_id)
        await page.keyboard.press("Enter")

        await asyncio.sleep(3) 
        target_link = page.locator(".product-name-link").first
        
        try:
            await target_link.wait_for(state="attached", timeout=8000)
        except:
            return "查无此产品", "搜索结果列表为空"

        async with context.expect_page(timeout=10000) as new_page_info:
            await target_link.evaluate("el => el.click()")
        new_page = await new_page_info.value
        href = new_page.url
        await new_page.close()

        if not href:
            return "跳转异常", "无法提取产品详情链接"

        # 2. 跳转询价历史
        final_url = re.sub(r"tab=[^&]+", "tab=inquiryHistoryTabPane", href)
        if "tab=" not in final_url:
            final_url += ("&" if "?" in final_url else "?") + "tab=inquiryHistoryTabPane"
        
        await page.goto(final_url)
        await asyncio.sleep(4) 
        await auto_clean_modals(page)

        # 3. 数据校验
        has_history = await page.evaluate("""() => {
            if (document.querySelector('.ant-empty, .ant-empty-image, .no-data')) return false;
            const text = document.body.innerText;
            if (text.includes('暂无数据') || text.includes('没有找到符合条件的记录')) return false;
            return true;
        }""")

        if not has_history:
            return "无询价历史", "产品详情页判定无记录"

        # 4. 下载附件
        jump_icon = page.locator(".jump-link .new-tab-icon").first
        downloaded_file_path = None
        
        try:
            async with context.expect_page(timeout=15000) as task_page_info:
                await jump_icon.click(force=True)
            
            task_page = await task_page_info.value
            await task_page.wait_for_load_state("domcontentloaded")
            await asyncio.sleep(4) 
            
            found_dl = await task_page.evaluate("""() => {
                const nodes = Array.from(document.querySelectorAll('a, span, button'));
                const btn = nodes.find(n => n.innerText && n.innerText.trim() === '下载');
                if (btn) { btn.id = 'final-download-node'; return true; }
                return false;
            }""")

            if found_dl:
                async with task_page.expect_download(timeout=30000) as download_info:
                    await task_page.locator("#final-download-node").click(force=True)
                download = await download_info.value
                original_filename = download.suggested_filename
                downloaded_file_path = os.path.join(DOWNLOAD_DIR, original_filename)
                await download.save_as(downloaded_file_path)
            else:
                await task_page.close()
                return "无Excel附件", "任务页无下载按钮"
            
            await task_page.close()
        except Exception as task_err:
            return "下载失败", f"抓取任务页失败: {str(task_err)[:50]}"

        # 5. 回传附件
        if downloaded_file_path and os.path.exists(downloaded_file_path):
            attachment_url = final_url.replace("tab=inquiryHistoryTabPane", "tab=attachmentTabPane")
            await page.goto(attachment_url)
            await asyncio.sleep(3) 
            
            try:
                upload_btn = page.locator("button").filter(has_text="上 传").first
                await upload_btn.click(force=True)
                await asyncio.sleep(1)
                
                local_upload_btn = page.locator("button").filter(has_text="上传本地文件").first
                async with page.expect_file_chooser(timeout=10000) as fc_info:
                    await local_upload_btn.click(force=True)
                
                file_chooser = await fc_info.value
                await file_chooser.set_files(downloaded_file_path)
                
                # 等待上传完成
                await asyncio.sleep(8)
                return "处理成功", f"文件 {os.path.basename(downloaded_file_path)} 回传完毕"
                
            except Exception as upload_err:
                return "上传失败", f"注入本地文件失败: {str(upload_err)[:50]}"
        else:
            return "文件丢失", "已触发下载但在本地未找到文件"

    except Exception as e:
        return "执行异常", f"未知错误: {str(e)[:50]}"


async def main():
    print("===================================================")
    print("      小满 CRM 自动化批处理系统 (v3.0)             ")
    print("===================================================")
    
    # 1. 检查任务文件
    if not os.path.exists(TASK_FILE):
        with open(TASK_FILE, 'w', encoding='utf-8') as f:
            f.write("HBB6969\nHBB3721\n")
        print(f"[!] 未发现任务清单。已在当前目录生成 {TASK_FILE}。")
        print(f">>> 请在 tasks.txt 中填入需要处理的产品编号（每行一个），然后重新运行。")
        return

    with open(TASK_FILE, 'r', encoding='utf-8') as f:
        tasks = [line.strip() for line in f if line.strip()]
    
    if not tasks:
        print("[!] tasks.txt 为空，请填入数据后运行。")
        return

    print(f"[*] 读取到 {len(tasks)} 个待处理产品编号。")
    
    # 初始化报告数据结构
    report_data = []
    report_filename = f"执行报告_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    report_filepath = os.path.join(REPORT_DIR, report_filename)

    # 2. 启动浏览器（全局复用）
    async with async_playwright() as p:
        print(f"[*] 启动 Edge 自动化实例 (强制虚拟分辨率)...")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_PATH, 
            executable_path=EDGE_PATH,
            headless=True,  # 保持静默运行
            # ========= 核心修改 =========
            no_viewport=False, # 关闭无视口模式
            viewport={"width": 1920, "height": 1080}, # 强制分配一个 1080P 的虚拟大屏
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0", # 伪装真实浏览器
            # ============================
            accept_downloads=True,
            args=[
                "--no-first-run", 
                "--disable-blink-features=AutomationControlled"
            ]
        )
        page = context.pages[0]
                
        try:
            # 3. 容灾：检查登录
            await check_login_status(page)
            
            # 4. 批处理循环
            for idx, product_id in enumerate(tasks, 1):
                print(f"\n>>> 进度: [{idx}/{len(tasks)}]")
                status, msg = await process_single_task(context, page, product_id)
                
                print(f"[{product_id}] 状态: {status} | 详情: {msg}")
                
                # 记录到报告
                report_data.append({
                    "处理时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "产品编号": product_id,
                    "执行状态": status,
                    "详细信息": msg
                })

        except KeyboardInterrupt:
            print("\n[!] 用户强制中止了批处理任务。")
        except Exception as global_err:
            print(f"\n[💥] 发生严重系统级错误: {global_err}")
        finally:
            await context.close()

    # 5. 生成巡检报告
    print("\n===================================================")
    print("[*] 正在生成自动化巡检报告...")
    with open(report_filepath, 'w', newline='', encoding='utf-8-sig') as csvfile:
        fieldnames = ['处理时间', '产品编号', '执行状态', '详细信息']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for data in report_data:
            writer.writerow(data)
            
    print(f"[SUCCESS] 批处理完成！报告已保存至:")
    print(f"   => {report_filepath}")
    print("===================================================")

if __name__ == "__main__":
    asyncio.run(main())