"""
결재하기 버튼 자동 클릭 스크립트

사용법:
    python apply_btn.py                          # 기본 경로 사용
    python apply_btn.py "경로\apply_btn.png"     # 이미지 경로 지정
    python apply_btn.py "경로\apply_btn.png" 0.9 # 이미지 경로 + confidence 지정
"""

import os
import sys
import time
from typing import Optional
import pyautogui as pag

# 기본 설정
DEFAULT_IMG_PATH = r"C:\Users\Code\mapping\apply_btn.png"
DEFAULT_CONFIDENCE = 0.8
DEFAULT_TIMEOUT = 10.0  # 초
PRE_CLICK_DELAY = 1.0  # 클릭 전 대기 시간 (초)
RETRY_INTERVAL = 0.5  # 재시도 간격 (초)

# OpenCV 설치 여부 확인
try:
    import cv2  # type: ignore
    OPENCV_AVAILABLE = True
except ImportError:
    OPENCV_AVAILABLE = False


def validate_image_file(image_path: str) -> None:
    """
    이미지 파일이 존재하고 읽을 수 있는지 확인합니다.
    
    Args:
        image_path: 확인할 이미지 파일 경로
    
    Raises:
        FileNotFoundError: 파일이 없을 때
        OSError: 파일을 읽을 수 없을 때
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(
            f"이미지 파일을 찾을 수 없습니다: {image_path}\n"
            f"파일 경로를 확인해주세요."
        )
    
    # 파일이 실제로 읽을 수 있는지 확인
    try:
        with open(image_path, 'rb') as f:
            f.read(1)  # 최소한 1바이트 읽기 시도
    except PermissionError:
        raise OSError(
            f"파일 읽기 권한이 없습니다: {image_path}\n"
            f"파일 권한을 확인해주세요."
        )
    except Exception as e:
        raise OSError(
            f"파일을 읽을 수 없습니다: {image_path}\n"
            f"오류: {e}\n"
            f"파일이 손상되었거나 지원하지 않는 형식일 수 있습니다."
        )


def find_button(image_path: str, confidence: Optional[float], timeout: float) -> tuple[int, int]:
    """
    화면에서 버튼 이미지를 찾아 좌표를 반환합니다.
    
    Args:
        image_path: 찾을 버튼 이미지 경로
        confidence: 이미지 매칭 신뢰도 (0.0 ~ 1.0), None이면 사용 안 함
        timeout: 최대 대기 시간 (초)
    
    Returns:
        (x, y) 좌표 튜플
    
    Raises:
        FileNotFoundError: 이미지 파일이 없을 때
        OSError: 파일을 읽을 수 없을 때
        TimeoutError: 타임아웃 내에 버튼을 찾지 못했을 때
    """
    # 파일 유효성 검사
    validate_image_file(image_path)
    
    start_time = time.time()
    while True:
        try:
            # OpenCV가 있고 confidence가 지정된 경우에만 confidence 사용
            if OPENCV_AVAILABLE and confidence is not None:
                pos = pag.locateCenterOnScreen(image_path, confidence=confidence)
            else:
                pos = pag.locateCenterOnScreen(image_path)
            
            if pos:
                return pos
        except pag.ImageNotFoundException:
            pass
        except OSError as e:
            # 파일 읽기 오류는 즉시 재발생
            raise OSError(
                f"이미지 파일을 읽을 수 없습니다: {image_path}\n"
                f"오류: {e}\n"
                f"해결 방법:\n"
                f"  - 파일 경로에 한글이 있으면 영문 경로로 복사해보세요\n"
                f"  - 파일이 손상되지 않았는지 확인하세요\n"
                f"  - 파일 형식이 PNG, JPG 등 지원 형식인지 확인하세요"
            )
        except NotImplementedError as e:
            # OpenCV가 없는데 confidence를 사용하려고 한 경우
            if "confidence" in str(e).lower():
                print("[경고] OpenCV가 설치되지 않아 confidence를 사용할 수 없습니다.")
                print("      confidence 없이 재시도합니다...")
                confidence = None  # confidence 비활성화
                continue
            raise
        
        elapsed = time.time() - start_time
        if elapsed >= timeout:
            confidence_msg = f" (confidence: {confidence})" if confidence else ""
            raise TimeoutError(
                f"타임아웃 ({timeout}초): 버튼 이미지를 찾을 수 없습니다.{confidence_msg}\n"
                f"해결 방법:\n"
                f"  - 줌 레벨 확인 (100% 권장)\n"
                f"  - 해상도 확인\n"
                f"  - 이미지 파일 다시 캡처\n"
                + (f"  - confidence 값 낮추기 (현재: {confidence})" if confidence else "  - OpenCV 설치: pip install opencv-python")
            )
        
        time.sleep(RETRY_INTERVAL)


def main() -> int:
    """메인 함수"""
    try:
        # 명령줄 인자 처리
        image_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_IMG_PATH
        confidence = float(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_CONFIDENCE
        
        if not (0.0 < confidence <= 1.0):
            print(f"[오류] confidence 값은 0.0 ~ 1.0 사이여야 합니다. (입력: {confidence})")
            return 1
        
        print(f"이미지 경로: {image_path}")
        if OPENCV_AVAILABLE:
            print(f"신뢰도: {confidence}")
        else:
            print(f"신뢰도: {confidence} (OpenCV 미설치로 사용 안 함)")
            print("      OpenCV 설치: pip install opencv-python")
            confidence = None  # OpenCV가 없으면 confidence 사용 안 함
        print(f"타임아웃: {DEFAULT_TIMEOUT}초")
        print(f"\n{PRE_CLICK_DELAY}초 후 버튼을 찾기 시작합니다...")
        print("(이 시간 동안 크롬 창을 활성화해주세요)")
        
        time.sleep(PRE_CLICK_DELAY)
        
        print("버튼 찾는 중...")
        pos = find_button(image_path, confidence, DEFAULT_TIMEOUT)
        
        print(f"버튼 발견! 위치: ({pos.x}, {pos.y})")
        pag.click(pos)
        print("✅ 결재하기 클릭 완료")
        
        return 0
        
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1
    except OSError as e:
        print(f"[오류] {e}")
        return 1
    except TimeoutError as e:
        print(f"[오류] {e}")
        return 1
    except KeyboardInterrupt:
        print("\n[중단] 사용자에 의해 취소되었습니다.")
        return 130
    except Exception as e:
        print(f"[오류] 예상치 못한 오류가 발생했습니다: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
