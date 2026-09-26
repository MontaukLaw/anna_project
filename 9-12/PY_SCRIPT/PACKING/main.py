"""自动装箱生成系统：程序入口。"""
from ui.app import PackingApp


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 3 and sys.argv[1] == '--self-test':
        from services.release_check import run_release_check
        raise SystemExit(run_release_check(sys.argv[2]))
    app = PackingApp()
    app.mainloop()
