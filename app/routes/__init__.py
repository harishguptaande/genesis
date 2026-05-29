from fastapi import APIRouter

from app.routes import items

router = APIRouter()
router.include_router(items.router, prefix="/items", tags=["items"])

__all__ = ["router"]
