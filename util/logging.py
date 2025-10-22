import logging
import sys

def setup_logger(
    name: str = None,
    level: int = logging.INFO,
    fmt: str = "[%(asctime)s]✅ %(levelname)s - %(message)s",
):
    """
    로거 설정 함수
    - name: 로거 이름 (모듈 단위 구분용)
    - level: 로깅 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL)
    - fmt: 출력 포맷 지정

    반환값:
        logging.Logger 객체
    """
    logger = logging.getLogger(name)

    # 동일 로거에 중복 핸들러가 쌓이지 않도록 초기화
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(fmt)
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(level)
    
    return logger