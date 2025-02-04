import yaml

from scraper import Scraper

CONFIG_PATH = "config.yaml"


def load_config(config_path: str) -> dict:
    with open(config_path, 'r', encoding='utf-8') as file:
        config = yaml.safe_load(file)

    return config


def main() -> None:
    config = load_config(CONFIG_PATH)

    scraper = Scraper(start_page_num=config['start_page_num'],
                      min_wait_time=config['min_wait_time'],
                      max_wait_time=config['max_wait_time'],
                      id2country=config['id2country'],
                      headers=config['headers'],
                      xlsx_path=config['xlsx_path'],
                      xlsx_page_title=config['xlsx_page_title'])
    scraper.scrape()


if __name__ == "__main__":
    main()
