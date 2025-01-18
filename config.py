import pdfkit

class PdfKitConfig:
    def __init__(self, path_to_wkhtmltopdf: str):
        """
        Initialize the PdfKitConfig class with the path to wkhtmltopdf.
        
        :param path_to_wkhtmltopdf: Full path to the wkhtmltopdf executable.
        """
        self.path_to_wkhtmltopdf = path_to_wkhtmltopdf
        self.config = self._create_configuration()

    def _create_configuration(self):
        """
        Creates and returns the configuration for pdfkit using the provided wkhtmltopdf path.
        
        :return: pdfkit configuration object.
        """
        return pdfkit.configuration(wkhtmltopdf=self.path_to_wkhtmltopdf)

    def get_configuration(self):
        """
        Returns the pdfkit configuration.
        
        :return: pdfkit configuration object.
        """
        return self.config
