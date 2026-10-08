from django.core.checks import Error, register
from django.db import connection
from django.db.migrations.recorder import MigrationRecorder
from django.db.utils import OperationalError, ProgrammingError
from services.models import Feature, FEATURE_CODES

@register()
def check_core_features_exist(app_configs, **kwargs):
    errors = []
    try:
        # 1. Check if migrations table exists
        recorder = MigrationRecorder(connection)
        if not recorder.has_table():
            return errors
            
        # 2. Check if the seed migration has been applied
        applied = recorder.applied_migrations()
        if ('services', '0003_seed_base_short_url_feature') not in applied:
            return errors
            
        # 3. Check if the feature table exists
        if 'services_feature' not in connection.introspection.table_names():
            return errors

        # 4. Now safe to check feature integrity
        db_features = set(Feature.objects.values_list('key', flat=True))
        for code in FEATURE_CODES:
            if code not in db_features:
                errors.append(
                    Error(
                        f"Core feature '{code}' is missing from the database.",
                        hint=f"Ensure migrations have run, or re-seed the feature.",
                        obj=Feature,
                        id='services.E001',
                    )
                )
    except (OperationalError, ProgrammingError):
        # Failsafe for missing tables or unavailable DB
        pass
    return errors
