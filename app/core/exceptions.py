class ReportEngineException(Exception):
    """Base exception for all report engine errors."""
    def __init__(self, message: str, status_code: int = 500):
        super().__init__(message)
        self.message = message
        self.status_code = status_code

class TemplateNotFoundException(ReportEngineException):
    def __init__(self, template_code: str):
        super().__init__(f"Template '{template_code}' not found in repository.", status_code=404)

class TemplateSyntaxException(ReportEngineException):
    def __init__(self, message: str):
        super().__init__(f"Template syntax error: {message}", status_code=400)

class RenderingException(ReportEngineException):
    def __init__(self, message: str):
        super().__init__(f"Rendering error: {message}", status_code=500)
