using System;
using Nop.Core;
using Nop.Core.Data;

using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using Nop.Core.Domain.Catalog;
using Nop.Data;


namespace Nop.Ncc.Repository
{
    //public interface INccRepository<T> : IRepository<T> where T : BaseEntity
    //{

    //    void Update(IEnumerable<T> entities);
    //}

    //public class NccRepository<T> : EfRepository<T>, INccRepository<T> where T : BaseEntity
    //{


    //    public NccRepository(IDbContext context) : base(context)
    //    {
    //    }

    //    public void Update(IEnumerable<T> entities)
    //    {
    //        try
    //        {
    //            if (entities == null)
    //                throw new ArgumentNullException("entity");

    //            this._context.SaveChanges();
    //        }
    //        catch (System.Data.Entity.Validation.DbEntityValidationException dbEx)
    //        {
    //            var msg = string.Empty;

    //            foreach (var validationErrors in dbEx.EntityValidationErrors)
    //                foreach (var validationError in validationErrors.ValidationErrors)
    //                    msg += Environment.NewLine + string.Format("Property: {0} Error: {1}", validationError.PropertyName, validationError.ErrorMessage);

    //            var fail = new Exception(msg, dbEx);
    //            //Debug.WriteLine(fail.Message, fail);
    //            throw fail;
    //        }
    //    }
    //}

    public static class NccRepositoryExtension
    {

        public static void Update(this EfRepository<Product> rep, IEnumerable<Product> entities) 
        {
            try
            {
                if (entities == null)
                    throw new ArgumentNullException("entity");

                rep.Context.SaveChanges();
            }
            catch (System.Data.Entity.Validation.DbEntityValidationException dbEx)
            {
                var msg = string.Empty;

                foreach (var validationErrors in dbEx.EntityValidationErrors)
                    foreach (var validationError in validationErrors.ValidationErrors)
                        msg += Environment.NewLine + string.Format("Property: {0} Error: {1}", validationError.PropertyName, validationError.ErrorMessage);

                var fail = new Exception(msg, dbEx);
                //Debug.WriteLine(fail.Message, fail);
                throw fail;
            }
        }
    }
}