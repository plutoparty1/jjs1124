import subprocess
import time
import sys

def run(cmd: str) -> int:
    return subprocess.call(cmd, shell=True)

def taskkill_graceful() -> int:
    # 정상 종료 시도: 실행 중인 엑셀에 종료 신호
    return run('taskkill /IM EXCEL.EXE')

def taskkill_force() -> int:
    # 강제 종료
    return run('taskkill /F /IM EXCEL.EXE')

def excel_running() -> bool:
    # EXCEL.EXE가 떠있는지 확인
    p = subprocess.run('tasklist /FI "IMAGENAME eq EXCEL.EXE"', shell=True, capture_output=True, text=True)
    return "EXCEL.EXE" in (p.stdout or "")

def main(timeout_sec: int = 5):
    if not excel_running():
        print("Excel not running.")
        return 0

    taskkill_graceful()

    # 정상 종료 기다림
    end = time.time() + timeout_sec
    while time.time() < end:
        if not excel_running():
            print("Excel closed gracefully.")
            return 0
        time.sleep(0.2)

    # 타임아웃이면 강제종료
    taskkill_force()
    time.sleep(0.5)

    if excel_running():
        print("Failed to kill Excel.")
        return 2

    print("Excel force-killed.")
    return 0

if __name__ == "__main__":
    t = 5
    if len(sys.argv) >= 2:
        try:
            t = int(sys.argv[1])
        except:
            pass
    raise SystemExit(main(t))
