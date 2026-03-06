from abc import ABC, abstractmethod


class CompanyAPI(ABC):

    # @abstractmethod
    # def steps(self):
    #     pass

    def add_awb(self, flight, carrier, manifest_path):
        raise NotImplementedError("This company does not support add_awb")

    def add_arn(self):
        raise NotImplementedError("This company does not support add_arn")

    def add_ens(self):
        raise NotImplementedError("This company does not support add_ens")

    def get_dsk(self):
        raise NotImplementedError("This company does not support get_dsk")

    def init_items_clearances(self):
        raise NotImplementedError("This company does not support init_items_clearances")

    def add_items_clearances(self):
        raise NotImplementedError("This company does not support add_items_clearances")

    def add_items_releases(self, flight, airway_bill, process_file_path):
        raise NotImplementedError("This company does not support add_items_releases")

    def add_items_releases_easy(self):
        raise NotImplementedError("This company does not support add_items_releases")

    def download_zc415(self, airway_bill):
        raise NotImplementedError("This company does not support download_zc415")

    def mark_arrived(self):
        raise NotImplementedError("This company does not support mark_arrived")