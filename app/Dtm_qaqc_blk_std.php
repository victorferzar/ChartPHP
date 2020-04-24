<?php

namespace App;

use Illuminate\Database\Eloquent\Model;

class Dtm_qaqc_blk_std extends Model
{
    //se establece el nombre de la tabla explicitamente para evitar 'serpent case'
    protected $table = 'DTM_QAQC_BLK_STD';
    //protected $dateFormat = 'U';
   /* protected $casts=[
      'RETURNDATE' =>'date'
    ];*/

}
