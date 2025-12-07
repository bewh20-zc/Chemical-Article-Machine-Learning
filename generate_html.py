import pandas as pd
import urllib.parse
import os
from datetime import datetime

def generate_search_page():
    """生成包含所有文献搜索链接的HTML页面"""
    try:
        # 读取Excel文件
        print("正在读取Excel文件...")
        df = pd.read_excel('table-721d94fb-eb22-4de8-a27a-f70d6650ae79.xlsx')
        print(f"成功读取 {len(df)} 条文献记录")
        
        # 统计信息
        total_papers = len(df)
        journal_count = df[df['文献/专利类型'] == '期刊论文'].shape[0] if '文献/专利类型' in df.columns else 0
        
        # 开始生成HTML
        print("正在生成HTML页面...")
        
        html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>文献批量检索工具 - {total_papers}篇文献</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
        }}
        .header {{
            background: #f0f0f0;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 20px;
        }}
        .paper {{
            border: 1px solid #ddd;
            padding: 15px;
            margin: 10px 0;
            border-radius: 5px;
            background: white;
        }}
        .paper:hover {{
            background: #f9f9f9;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}
        .paper-title {{
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 5px;
            font-size: 16px;
        }}
        .paper-meta {{
            color: #7f8c8d;
            font-size: 14px;
            margin-bottom: 10px;
        }}
        .search-btn {{
            display: inline-block;
            margin: 5px 10px 5px 0;
            padding: 8px 15px;
            background: #4285f4;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            font-size: 14px;
        }}
        .search-btn:hover {{
            background: #3367d6;
        }}
        .search-btn.sciencedirect {{
            background: #ff6b35;
        }}
        .search-btn.researchgate {{
            background: #00ccbb;
        }}
        .progress {{
            margin: 20px 0;
            padding: 10px;
            background: #e8f4fd;
            border-radius: 5px;
        }}
        .controls {{
            margin: 20px 0;
        }}
        .control-btn {{
            padding: 10px 20px;
            margin: 0 10px 10px 0;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }}
        .paper-number {{
            float: right;
            background: #4285f4;
            color: white;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 12px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📚 文献批量检索工具</h1>
        <p>共找到 <strong>{total_papers}</strong> 篇文献，点击下方按钮可一键搜索</p>
        <p>生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>
    
    <div class="controls">
        <button class="control-btn" onclick="openAllGoogle()">一键打开所有Google链接</button>
        <button class="control-btn" onclick="markAllDone()">标记所有为已下载</button>
        <button class="control-btn" onclick="showUndoneOnly()">只显示未下载</button>
    </div>
    
    <div class="progress">
        处理进度：<span id="progress">0/{total_papers}</span>
        <div style="background: #ddd; height: 10px; border-radius: 5px; margin-top: 5px;">
            <div id="progressBar" style="background: #4CAF50; height: 100%; width: 0%; border-radius: 5px;"></div>
        </div>
    </div>
    
    <div id="paperList">
"""

        # 生成每个文献的条目
        for index, row in df.iterrows():
            if pd.notna(row.get('文献/专利名称')):
                paper_num = index + 1
                title = str(row['文献/专利名称']).strip()
                authors = str(row.get('参考文献条目', ''))[:80]
                year = str(row.get('年份', '')) if pd.notna(row.get('年份')) else ''
                journal = str(row.get('期刊名称/专利号', '')) if pd.notna(row.get('期刊名称/专利号')) else ''
                
                # 对搜索词进行编码
                query = urllib.parse.quote(title)
                
                html_content += f"""
        <div class="paper" id="paper{paper_num}">
            <div class="paper-number">#{paper_num}</div>
            <div class="paper-title">{title}</div>
            <div class="paper-meta">
                作者：{authors}<br>
                年份：{year} | 期刊：{journal}
            </div>
            <div>
                <a href="https://scholar.google.com/scholar?q={query}" target="_blank" class="search-btn" onclick="markDone({paper_num})">
                    🔍 Google Scholar
                </a>
                <a href="https://www.sciencedirect.com/search?qs={query}" target="_blank" class="search-btn sciencedirect" onclick="markDone({paper_num})">
                    📰 ScienceDirect
                </a>
                <a href="https://www.researchgate.net/search/publication?q={query}" target="_blank" class="search-btn researchgate" onclick="markDone({paper_num})">
                    👥 ResearchGate
                </a>
            </div>
            <div style="margin-top: 10px;">
                <input type="checkbox" id="check{paper_num}" onchange="togglePaper({paper_num})">
                <label for="check{paper_num}">已下载</label>
            </div>
        </div>
"""

        # 添加JavaScript和HTML结尾
        html_content += f"""
    </div>
    
    <script>
        // 保存已处理的文献编号
        let donePapers = JSON.parse(localStorage.getItem('donePapers')) || [];
        
        // 页面加载时恢复状态
        window.onload = function() {{
            donePapers.forEach(num => {{
                const checkbox = document.getElementById('check' + num);
                const paper = document.getElementById('paper' + num);
                if (checkbox) checkbox.checked = true;
                if (paper) paper.style.opacity = '0.6';
            }});
            updateProgress();
        }};
        
        // 标记为已下载
        function markDone(num) {{
            if (!donePapers.includes(num)) {{
                donePapers.push(num);
                localStorage.setItem('donePapers', JSON.stringify(donePapers));
                
                const paper = document.getElementById('paper' + num);
                if (paper) paper.style.opacity = '0.6';
                
                const checkbox = document.getElementById('check' + num);
                if (checkbox) checkbox.checked = true;
                
                updateProgress();
            }}
        }}
        
        // 切换论文状态
        function togglePaper(num) {{
            const checkbox = document.getElementById('check' + num);
            const paper = document.getElementById('paper' + num);
            
            if (checkbox.checked) {{
                if (!donePapers.includes(num)) {{
                    donePapers.push(num);
                }}
                if (paper) paper.style.opacity = '0.6';
            }} else {{
                const index = donePapers.indexOf(num);
                if (index > -1) {{
                    donePapers.splice(index, 1);
                }}
                if (paper) paper.style.opacity = '1';
            }}
            
            localStorage.setItem('donePapers', JSON.stringify(donePapers));
            updateProgress();
        }}
        
        // 一键打开所有Google链接
        function openAllGoogle() {{
            const undone = [];
            for (let i = 1; i <= {total_papers}; i++) {{
                if (!donePapers.includes(i)) {{
                    undone.push(i);
                }}
            }}
            
            if (undone.length === 0) {{
                alert('所有文献都已处理！');
                return;
            }}
            
            if (confirm('将打开 ' + undone.length + ' 个未处理文献的搜索页面，继续吗？')) {{
                // 每次最多打开5个，避免浏览器崩溃
                const batchSize = 5;
                for (let i = 0; i < Math.min(batchSize, undone.length); i++) {{
                    const paperId = undone[i];
                    const link = document.querySelector('#paper' + paperId + ' a[href*="scholar.google.com"]');
                    if (link) {{
                        window.open(link.href, '_blank');
                        markDone(paperId);
                    }}
                }}
                
                if (undone.length > batchSize) {{
                    alert('已打开前 ' + batchSize + ' 个，剩下的请继续处理。');
                }}
            }}
        }}
        
        // 标记所有为已下载
        function markAllDone() {{
            if (confirm('标记所有文献为已下载吗？')) {{
                for (let i = 1; i <= {total_papers}; i++) {{
                    const checkbox = document.getElementById('check' + i);
                    const paper = document.getElementById('paper' + i);
                    if (checkbox) checkbox.checked = true;
                    if (paper) paper.style.opacity = '0.6';
                    
                    if (!donePapers.includes(i)) {{
                        donePapers.push(i);
                    }}
                }}
                localStorage.setItem('donePapers', JSON.stringify(donePapers));
                updateProgress();
                alert('已标记所有文献为已下载');
            }}
        }}
        
        // 只显示未下载的文献
        function showUndoneOnly() {{
            const allPapers = document.querySelectorAll('.paper');
            allPapers.forEach(paper => {{
                const paperId = parseInt(paper.id.replace('paper', ''));
                paper.style.display = donePapers.includes(paperId) ? 'none' : 'block';
            }});
            
            const undoneCount = {total_papers} - donePapers.length;
            alert('显示了 ' + undoneCount + ' 篇未下载的文献');
        }}
        
        // 更新进度显示
        function updateProgress() {{
            const doneCount = donePapers.length;
            const total = {total_papers};
            const percent = Math.round((doneCount / total) * 100);
            
            document.getElementById('progress').textContent = doneCount + '/' + total;
            document.getElementById('progressBar').style.width = percent + '%';
            
            // 更新页面标题显示进度
            document.title = '文献检索 (' + doneCount + '/' + total + ') - ' + percent + '%';
        }}
        
        // 快捷键支持
        document.addEventListener('keydown', function(e) {{
            // Ctrl+Shift+D 标记所有
            if (e.ctrlKey && e.shiftKey && e.key === 'D') {{
                markAllDone();
            }}
        }});
    </script>
    
    <div style="margin-top: 40px; padding: 20px; text-align: center; color: #666; border-top: 1px solid #eee;">
        <p>使用说明：</p>
        <ol style="text-align: left; display: inline-block; margin: 10px auto;">
            <li>点击任意搜索按钮会在新标签页打开文献搜索页面</li>
            <li>在搜索页面找到并下载文献PDF</li>
            <li>返回此页面，勾选"已下载"或页面会自动标记</li>
            <li>进度会自动保存，关闭浏览器后重新打开仍有效</li>
        </ol>
        <p style="margin-top: 20px;">© 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>
</body>
</html>"""
        
        # 保存HTML文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"文献检索工具_{timestamp}.html"
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ 成功生成HTML文件：{filename}")
        print(f"✅ 共包含 {total_papers} 篇文献")
        print(f"✅ 请用浏览器打开这个HTML文件开始检索")
        
        return filename
        
    except Exception as e:
        print(f"❌ 出错了：{str(e)}")
        print("请检查：")
        print("1. Excel文件是否在同一个文件夹中")
        print("2. Excel文件名是否正确")
        print("3. 是否安装了pandas库（运行：pip install pandas）")
        return None

# 主程序
if __name__ == "__main__":
    print("=" * 50)
    print("文献批量检索工具 v1.0")
    print("=" * 50)
    
    # 检查Excel文件是否存在
    excel_file = 'table-721d94fb-eb22-4de8-a27a-f70d6650ae79.xlsx'
    if not os.path.exists(excel_file):
        print(f"❌ 找不到Excel文件：{excel_file}")
        print("请确保Excel文件放在同一文件夹中")
        input("按回车键退出...")
    else:
        filename = generate_search_page()
        if filename:
            print("\n✨ 下一步操作：")
            print(f"1. 在文件夹中找到并双击打开：{filename}")
            print("2. 点击文献旁边的搜索按钮")
            print("3. 在新标签页中下载文献")
            print("4. 返回标记已下载的文献")
            print("\n💡 提示：进度会自动保存，下次打开还能继续")
            
            # 询问是否自动打开
            choice = input("\n是否立即用浏览器打开生成的HTML文件？(y/n): ")
            if choice.lower() == 'y':
                import webbrowser
                webbrowser.open(f'file://{os.path.abspath(filename)}')
                print("✅ 已在浏览器中打开！")
            
        input("\n按回车键退出程序...")