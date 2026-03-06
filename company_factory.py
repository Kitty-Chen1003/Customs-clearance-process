from ecargo_api import EcargoAPI
from lsas_api import LSASAPI

APIS = {
    "Ecargo": EcargoAPI,
    "LSAS": LSASAPI,
}


class CompanyFactory:
    @staticmethod
    def get_api(company_name: str):
        try:
            return APIS[company_name]()
        except KeyError:
            raise ValueError(f"Unsupported company: {company_name}")
