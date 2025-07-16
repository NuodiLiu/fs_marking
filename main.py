from config import test1_config
from core.pipeline import run_batch
from core.utils.utils import validate_config

def main():
    config = test1_config
    validate_config(config)
    run_batch(config)

if __name__ == "__main__":
    main()
