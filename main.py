import logging
from sap_gui import SapGui
from typing import Dict, Optional

def get_sap_config() -> Dict[str, str]:
    """Return SAP configuration parameters."""
    return {
        "platform": "SAP PRD",  # SAP system identifier
        "client": "110",  # Client number
        "username": "RPA_USER",
        "password": "Dezembro#2024",
        "language": "PT"  # Language code
    }

def create_sales_order(sap: SapGui) -> Optional[str]:
    """Example function to create a sales order in SAP."""
    try:
        # Navigate to VA01 transaction
        sap.perform_operation("/nVA01", "wnd[0]/usr/ctxtVBAK-AUART")

        input("Press Enter after VA01 transaction is open...")

        # Additional order creation logic...

        return "Order created successfully"
    except Exception as e:
        logging.error(f"Failed to create sales order: {str(e)}")
        return None

def main():
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    # Initialize SAP session
    try:
        sap = SapGui(get_sap_config())

        # Login to SAP
        if not sap.sapLogin():
            raise Exception("Failed to login to SAP")

        # Perform business operations
        result = create_sales_order(sap)
        if result:
            logging.info(result)

    except Exception as e:
        logging.error(f"SAP automation failed: {str(e)}")

    finally:
        # Ensure proper logout and cleanup
        try:
            sap.sapLogout()
            sap.close_connection()
        except Exception as e:
            logging.error(f"Cleanup failed: {str(e)}")

if __name__ == "__main__":
    main()
