using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  internal class NavigationCollection
  {
    ImportTypeVM importTypeVM = new ImportTypeVM();
    List<NavigationItem> timeSeriesNavigation = new List<NavigationItem>();
    List<NavigationItem> pairedDataNavigation = new List<NavigationItem>();
    MainViewModel model;
    public NavigationCollection(MainViewModel model)
    {
      this.model = model;
      var rootNavigation = new NavigationItem
      {
        ViewModel = null,
        UserControl = new ImportTypeView(model.ImportTypeVM),
        BackEnabled = false,
        NextEnabled = true,
      };

      CreateTimeSeriesNavagation(rootNavigation);
      CreatePairedDataNavagation(rootNavigation);
    }
    public NavigationItem this[int index]
    {
      get { 
        if( model.ImportTypeVM.SelectedImportType.Type == ImportType.TimeSeries)
        return timeSeriesNavigation[index]; 
      else
          return pairedDataNavigation[index];
      }
    }


    private void CreatePairedDataNavagation(NavigationItem rootNavigation)
    {
      pairedDataNavigation.Add(rootNavigation);

      RangeSelectionVM vm = new RangeSelectionPairedDataX(model);

      pairedDataNavigation.Add(new NavigationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

      vm = new RangeSelectionPairedDataY(model);

      pairedDataNavigation.Add(new NavigationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

      var pdReview = new PairedDataReviewView(model);
      var reviewVM = new PairedDataReviewVM(pdReview.WorkSheet, model.DssFileName);
      pairedDataNavigation.Add(new NavigationItem
      {
        ViewModel = reviewVM,
        UserControl = pdReview,
        BackEnabled = true,
        NextEnabled = true,
        FinalStep= true
      });




    }

    private void CreateTimeSeriesNavagation(NavigationItem rootNavigation)
    {
      timeSeriesNavigation.Add(rootNavigation);

      RangeSelectionVM vm = new RangeSelectionDatesVM(model);
      timeSeriesNavigation.Add(new NavigationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

      vm = new RangeSelectionTimeSeriesValues(model);
      timeSeriesNavigation.Add(new NavigationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });


      var reviewControl = new TimeSeriesReviewView(model);
      var reviewVM = new TimeSeriesReviewVM(reviewControl.WorkSheet, model.DssFileName);
      timeSeriesNavigation.Add(new NavigationItem
      {
        ViewModel = reviewVM,
        UserControl = reviewControl,
        BackEnabled = true,
        NextEnabled = true,
        FinalStep = true,
      });
    }


  }
}
