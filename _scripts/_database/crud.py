from datetime import datetime, date
from sqlalchemy.orm import sessionmaker
from sqlalchemy.sql import update
from . import models, schemas, utils
import json

async def create_purchase_request(db: sessionmaker, content: schemas.PurchaseRequestBase):
    db_purchase = models.PurchaseRequests(
            requester=content.requester, budget_id=content.budget_id,
            provider_id=content.provider_id, _type=content._type,
            payment_terms=content.payment_terms, items=content.items,
            )
    db.add(db_purchase)
    db.commit()
    db.refresh(db_purchase)
    return db_purchase

# ################################################################################
# # READ
# async def read_user(db: Session, by: str, parameter: str|int|float|date):
#     match by:
#         case 'id':
#             return db.query(models.Users).filter(models.Users.id==parameter).all()
#         case 'email':
#             return db.query(models.Users).filter(models.Users.email==parameter).all()
#         case 'name':
#             return db.query(models.Users).filter(models.Users.name==parameter).all()
#         case 'tag':
#             return db.query(models.Users).filter(models.Users.tag==parameter).all()
#         case 'status':
#             return db.query(models.Users).filter(models.Users.status==parameter).all()
#         case _:
#             return []
#
def get_accredited_person_by_id(db: sessionmaker, id: str):
    return db.query().filter(models.AccreditedPersons.id==id).first()
def get_purchase_request_by_id(db: sessionmaker, id: int):
    return db.query(models.PurchaseRequests).filter(models.PurchaseRequests.id==id).first()

#
# ################################################################################
# # UPSERT
# async def update_user(db: Session, content: schemas.UserUpdate):
#     content_dict = content.dict()
#     if content_dict['password']:
#         content_dict['hashed_password'] = tools.encrypt_pass(content_dict['password'])
#     del content_dict['confirming_password'], content_dict['id'], content_dict['password']
#     content_dict['updated_at'] = datetime.now()
#     content_dict = dict((k, v) for k, v in content_dict.items() if v is not None)
#     db_user = db.query(models.Users).filter(models.Users.id==content.id).update(content_dict, synchronize_session=False)
#     db.commit()
#     return db_user
#
# ################################################################################
# # DELETE
# async def delete_address(db: Session, content: schemas.AddressDelete):
#     db_address = db.query(models.Addressess).filter(models.Addressess.id==content.id).first()
#     db.delete(db_address)
#     db.commit()
#     return []
