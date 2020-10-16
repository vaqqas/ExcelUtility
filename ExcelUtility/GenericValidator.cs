using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;


namespace Vqs.Excel
{
    /// <summary>
    /// Fires the attribute based validaters on the model class even outside an MVC application.
    /// </summary>
    public class GenericValidator
    {
        /// <summary>
        /// Fire the validation on Model object. Doesn't throws exceptions if validation fails.
        /// </summary>
        /// <param name="obj">Model object</param>
        /// <param name="results">results of the validations fired on the model properties</param>
        /// <returns>true if all validations passed</returns>
        public static bool TryValidate(object obj, out ICollection<ValidationResult> results)
        {
            var context = new ValidationContext(obj, serviceProvider: null, items: null);
            results = new List<ValidationResult>();
            return Validator.TryValidateObject(
                obj, context, results,
                validateAllProperties: true
            );
        }

        /// <summary>
        /// Fire the validation on Model object. Doesn't throws exceptions if validation fails.
        /// </summary>
        /// <param name="obj">Model object</param>
        /// <param name="errorMessage">All Error messages</param>
        /// <returns>true if all validations passed</returns>
        public static bool TryValidateAny(object obj, out string errorMessage)
        {
            ICollection<ValidationResult> results;
            errorMessage = "";
            bool flag = TryValidate(obj, out results);
            if( !flag)
            {
                StringBuilder sb = new StringBuilder();
                foreach(ValidationResult result in results)
                {
                    sb.Append(result.ErrorMessage + "; ");
                }

                errorMessage = sb.ToString();
            }

            return flag;
        }

        /// <summary>
        /// Fire the validation on an BaseExcelModel Model object. Doesn't throws exceptions if validation fails.
        /// sets the values of the IsValid and ErrorMessage inside the BaseExcelModel object
        /// </summary>
        /// <param name="model">BaseExcelModel Model object</param>
        /// <returns>true if all validations passed</returns>
        public static bool TryValidate(object model)
        {
            ICollection<ValidationResult> results;
            string errorMessage = "";
            bool flag = TryValidate(model, out results);

            if (!flag)
            {
                StringBuilder sb = new StringBuilder();
                foreach (ValidationResult result in results)
                {
                    sb.Append(result.ErrorMessage + "; ");
                }
                errorMessage = sb.ToString();
            }

            BaseExcelModel emodel = model as BaseExcelModel;
            if (emodel != null)
            {
                emodel.ErrorMessage = errorMessage;
                emodel.IsValid = flag;
            }

            return flag;
        }

        /// <summary>
        /// Fire the validation on a collection of BaseExcelModel Model object. Doesn't throws exceptions if validation fails.
        /// sets the values of the IsValid and ErrorMessage inside each of the BaseExcelModel object in the collection
        /// </summary>
        /// <param name="models">Collection of the BaseExcelModel models</param>
        /// <returns>number of records failed validation</returns>
        public static int TryValidate(IEnumerable<object> models)
        {
            int failedcount = 0;
            Parallel.ForEach(models, (model) =>
            {
                if (!TryValidate(model))
                    failedcount++;
            });            

            return failedcount;
        }
    }
}
