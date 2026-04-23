using Thermo.SampleManager.Common.Data;
using Thermo.SampleManager.Library.EntityDefinition;

namespace Thermo.SampleManager.ObjectModel
{
	/// <summary>
	/// Defines extended business logic and manages access to the HAZARD entity.
	/// </summary> 
	[SampleManagerEntity( "HAZARD" )]
	public class ExtendedHazard : Hazard
	{
		/// <summary>
		/// Gets the LabelText2 property.
		/// </summary>
		/// <remarks>
		/// Demonstrates an example of extending the hazard with a virtual property based on
		/// manipulating a real property.
		/// </remarks>
		/// <value>The LabelText2.</value>
		[PromptText]
		public string LabelText2
		{
			get
			{
				if ( LabelText.Length >= 3 )
				{
					return LabelText.Substring( 0, 3 );
				}

				return "Too Short";
			}
		}

		[PromptInteger]
		public int Random
		{
			get
			{
				return Library.Utils.Random(1, 100);
			}
		}
	}
}
