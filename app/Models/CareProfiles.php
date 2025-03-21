<?php

namespace App\Models;

//use Illuminate\Database\Eloquent\Factories\HasFactory;

use DateTime;
use DateTimeZone;
use Exception;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CareProfiles extends Model
{
    use SoftDeletes;

    protected $table = 'tbl_care_profiles';

    public function careProfilesFoms() {
        return $this->belongsToMany(CareProfilesFoms::class, 'tbl_care_profile_care_profile_foms', 'care_profile_id', 'care_profile_foms_id');
    }

    /**
     * Возвращает соответствующий профиль из приказа №17 (name, code) на указанную дату (текущую, если не указана)
     */
    public function decreeN17(DateTime|null $dt = null) {
        if ($dt === null) {
            $dt = new DateTime('now');
            $dt->setTimezone(new DateTimeZone("UTC"));
        }
        $dtStr = $dt->format('Y-m-d');

        $d = CareProfilesDecreeN17::Where('care_profile_id', $this->id)->whereRaw("? BETWEEN effective_from AND effective_to", [$dtStr])->get();

        if ($d->count() !== 1) {
            throw new Exception("Профиль медицинской помощи ($this->id на $dtStr) должен соответствовать одному профилю из приказа №17");
        }

        return $d[0];
    }
}
