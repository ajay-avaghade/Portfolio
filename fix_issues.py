import os
import re

directory = "/Users/ajayavaghade/Portfolio/assets"
for filename in os.listdir(directory):
    if filename.startswith("build-") and filename.endswith(".js"):
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r') as f:
            content = f.read()
        
        # Fix bullet points
        content = re.sub(r'(y:\s*[0-9.]+,\s*w:\s*[0-9.]+)(,\s*fontSize:\s*[0-9]+[^}]+bullet:\s*true)', r'\1, h: 2.5, valign: "top"\2', content)
        
        # Fix sub-headers on Slide 2 and 3
        content = re.sub(r'(y:\s*1\.[5-7],\s*w:\s*[0-9.]+)(,\s*fontSize:\s*1[4-6],\s*bold:\s*true)', r'\1, h: 0.4, valign: "top"\2', content)
        
        # For titles of slides (y: 0.5)
        content = re.sub(r'(y:\s*0\.5,\s*w:\s*[0-9.]+)(,\s*fontSize:\s*20)', r'\1, h: 0.6, valign: "top"\2', content)
        
        with open(filepath, 'w') as f:
            f.write(content)
        print(f"Fixed {filename}")

# Fix prototypes
html_dir = "/Users/ajayavaghade/Portfolio"
js_snippet = """
<script>
document.querySelectorAll('.nav-item').forEach(el => {
    el.addEventListener('click', () => {
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        el.classList.add('active');
    });
});
document.querySelectorAll('button').forEach(btn => {
    if(!btn.classList.contains('filter-btn')) {
        btn.addEventListener('click', () => {
            const originalText = btn.innerText;
            btn.innerText = 'Processing...';
            btn.style.opacity = '0.7';
            setTimeout(() => {
                btn.innerText = 'Success \u2713';
                btn.style.background = '#10B981';
                btn.style.color = '#fff';
                btn.style.borderColor = '#10B981';
                btn.style.opacity = '1';
                setTimeout(() => {
                    btn.innerText = originalText;
                    btn.style.background = '';
                    btn.style.borderColor = '';
                }, 2000);
            }, 800);
        });
    }
});
document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.filter-btn').forEach(n => n.classList.remove('active'));
        btn.classList.add('active');
    });
});
</script>
</body>
"""

for filename in os.listdir(html_dir):
    if filename.endswith("-prototype.html"):
        filepath = os.path.join(html_dir, filename)
        with open(filepath, 'r') as f:
            content = f.read()
        
        if "<script>" not in content:
            content = content.replace("</body>", js_snippet)
            with open(filepath, 'w') as f:
                f.write(content)
            print(f"Fixed {filename}")
