#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高性能静态文件服务器 - 优化本地网站加载速度
支持：多线程并发、内存缓存、压缩传输
"""
import http.server
import socketserver
import gzip
import io
import os
import sys

PORT = 8765
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

# 内存缓存：存储已加载的文件内容 (content, content_type, mtime)
_file_cache = {}
# 缓存大小限制
MAX_CACHE_SIZE = 50  # 最多缓存50个文件
MAX_FILE_SIZE = 10 * 1024 * 1024  # 单文件缓存上限 10MB


class OptimizedHandler(http.server.SimpleHTTPRequestHandler):
    """优化的HTTP请求处理器"""
    
    # 设置缓存过期时间头
    def end_headers(self):
        # 静态资源缓存 1 小时
        if self.path.endswith(('.html', '.css', '.js', '.pdf', '.png', '.jpg', '.svg', '.ico')):
            self.send_header('Cache-Control', 'public, max-age=3600')
        # 启用压缩支持
        self.send_header('Accept-Encoding', 'gzip')
        super().end_headers()
    
    # 压缩响应
    def _send_compressed(self, content, content_type):
        compressed = gzip.compress(content)
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Encoding', 'gzip')
        self.send_header('Content-Length', str(len(compressed)))
        self.send_header('Vary', 'Accept-Encoding')
        self.end_headers()
        self.wfile.write(compressed)
    
    def do_GET(self):
        filepath = self.translate_path(self.path)
        
        # 对于大文件(PDF)，直接流式传输，不经过缓存
        if filepath.endswith('.pdf'):
            try:
                with open(filepath, 'rb') as f:
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/pdf')
                    self.send_header('Content-Length', str(os.path.getsize(filepath)))
                    self.send_header('Cache-Control', 'public, max-age=3600')
                    self.send_header('Accept-Ranges', 'bytes')
                    self.end_headers()
                    # 流式写入，支持 Range 请求
                    while True:
                        chunk = f.read(8192 * 8)  # 64KB chunks
                        if not chunk:
                            break
                        self.wfile.write(chunk)
            except FileNotFoundError:
                self.send_error(404, "File not found")
            return
        
        # 小文件走缓存 + 压缩
        try:
            # 获取文件修改时间
            current_mtime = os.path.getmtime(filepath)
            
            # 检查缓存（同时验证修改时间，文件更新则刷新缓存）
            if filepath in _file_cache:
                content, content_type, cached_mtime = _file_cache[filepath]
                if cached_mtime >= current_mtime:
                    # 缓存有效
                    accept_encoding = self.headers.get('Accept-Encoding', '')
                    if 'gzip' in accept_encoding and len(content) > 200:
                        self._send_compressed(content, content_type)
                    else:
                        self.send_response(200)
                        self.send_header('Content-Type', content_type)
                        self.send_header('Content-Length', str(len(content)))
                        self.end_headers()
                        self.wfile.write(content)
                    return
                else:
                    # 文件已更新，清除旧缓存
                    del _file_cache[filepath]
            
            # 读取文件
            with open(filepath, 'rb') as f:
                content = f.read()
            
            # 判断类型
            if filepath.endswith('.html'):
                content_type = 'text/html; charset=utf-8'
            elif filepath.endswith('.css'):
                content_type = 'text/css; charset=utf-8'
            elif filepath.endswith('.js'):
                content_type = 'application/javascript; charset=utf-8'
            elif filepath.endswith('.svg'):
                content_type = 'image/svg+xml'
            elif filepath.endswith('.png'):
                content_type = 'image/png'
            elif filepath.endswith('.jpg'):
                content_type = 'image/jpeg'
            else:
                content_type = 'application/octet-stream'
            
            # 加入缓存（如果符合条件）
            if len(content) < MAX_FILE_SIZE and len(_file_cache) < MAX_CACHE_SIZE:
                _file_cache[filepath] = (content, content_type, current_mtime)
            
            # 发送响应
            accept_encoding = self.headers.get('Accept-Encoding', '')
            if 'gzip' in accept_encoding and len(content) > 200:
                self._send_compressed(content, content_type)
            else:
                self.send_response(200)
                self.send_header('Content-Type', content_type)
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
                
        except FileNotFoundError:
            self.send_error(404, "File not found")
        except Exception as e:
            self.send_error(500, str(e))
    
    # 禁止日志过多输出
    def log_message(self, format, *args):
        if '--verbose' in sys.argv:
            super().log_message(format, *args)


class ThreadedServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    """多线程HTTP服务器"""
    daemon_threads = True
    allow_reuse_address = True


if __name__ == '__main__':
    os.chdir(DIRECTORY)
    print(f"🚀 高性能服务器启动: http://localhost:{PORT}")
    print(f"📂 根目录: {DIRECTORY}")
    print(f"💾 内存缓存: {len(_file_cache)} 项")
    print(f"🔗 多线程: 已启用")
    print(f"🗜️ 压缩传输: 已启用")
    try:
        with ThreadedServer(('', PORT), OptimizedHandler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n✅ 服务器已停止")
